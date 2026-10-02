using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxRenderPreprocessor
    {
        public static bool NeedsAlternateContentFallback(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return false;

            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension != ".pptx" && extension != ".pptm")
                return false;

            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(path))
                {
                    for (int i = 0; i < archive.Entries.Count; i++)
                    {
                        ZipArchiveEntry entry = archive.Entries[i];
                        string name = entry.FullName.Replace('\\', '/');

                        if (!IsPresentationXml(name))
                            continue;

                        using (Stream stream = entry.Open())
                        using (StreamReader reader = new StreamReader(
                            stream,
                            Encoding.UTF8,
                            true,
                            4096,
                            false))
                        {
                            string text = reader.ReadToEnd();
                            if (text.IndexOf(
                                    "AlternateContent",
                                    StringComparison.Ordinal) >= 0)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        public static string PrepareForRendering(string sourcePath)
        {
            if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                throw new FileNotFoundException("PPTX file was not found.", sourcePath);

            string extension = Path.GetExtension(sourcePath).ToLowerInvariant();
            if (extension != ".pptx" && extension != ".pptm")
                return sourcePath;

            if (!NeedsAlternateContentFallback(sourcePath))
                return sourcePath;

            string cacheRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PowerPointLite",
                "RenderPrep");
            Directory.CreateDirectory(cacheRoot);

            string cacheKey = BuildCacheKey(sourcePath);
            string destination = Path.Combine(
                cacheRoot,
                cacheKey + extension);

            if (File.Exists(destination))
                return destination;

            string stage = destination + ".writing";
            DeleteIfExists(stage);

            try
            {
                using (ZipArchive input = ZipFile.OpenRead(sourcePath))
                using (ZipArchive output = ZipFile.Open(stage, ZipArchiveMode.Create))
                {
                    for (int i = 0; i < input.Entries.Count; i++)
                    {
                        ZipArchiveEntry source = input.Entries[i];
                        ZipArchiveEntry target = output.CreateEntry(
                            source.FullName,
                            CompressionLevel.Optimal);

                        if (string.IsNullOrEmpty(source.Name))
                            continue;

                        if (IsPresentationXml(source.FullName))
                        {
                            CopyPresentationXml(source, target);
                        }
                        else
                        {
                            using (Stream sourceStream = source.Open())
                            using (Stream targetStream = target.Open())
                                sourceStream.CopyTo(targetStream);
                        }
                    }
                }

                if (File.Exists(destination))
                    File.Delete(destination);

                File.Move(stage, destination);
                return destination;
            }
            catch
            {
                DeleteIfExists(stage);
                throw;
            }
        }

        private static void CopyPresentationXml(
            ZipArchiveEntry source,
            ZipArchiveEntry target)
        {
            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = false;

            using (Stream input = source.Open())
                document.Load(input);

            FlattenAlternateContent(document);

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;

            using (Stream output = target.Open())
            using (XmlWriter writer = XmlWriter.Create(output, settings))
                document.Save(writer);
        }

        private static void FlattenAlternateContent(XmlDocument document)
        {
            if (document == null || document.DocumentElement == null)
                return;

            List<XmlNode> nodes = new List<XmlNode>();
            CollectAlternateContent(document.DocumentElement, nodes);

            for (int i = nodes.Count - 1; i >= 0; i--)
            {
                XmlNode alternate = nodes[i];
                XmlNode parent = alternate.ParentNode;
                if (parent == null)
                    continue;

                XmlNode selected = DirectChild(alternate, "Fallback");

                if (selected == null)
                {
                    for (int c = 0; c < alternate.ChildNodes.Count; c++)
                    {
                        XmlNode candidate = alternate.ChildNodes[c];
                        if (candidate.LocalName == "Choice")
                        {
                            selected = candidate;
                            break;
                        }
                    }
                }

                if (selected != null)
                {
                    List<XmlNode> replacements = new List<XmlNode>();

                    for (int c = 0; c < selected.ChildNodes.Count; c++)
                    {
                        XmlNode child = selected.ChildNodes[c];
                        if (child.NodeType == XmlNodeType.Element ||
                            child.NodeType == XmlNodeType.Text ||
                            child.NodeType == XmlNodeType.CDATA)
                        {
                            replacements.Add(child);
                        }
                    }

                    for (int c = 0; c < replacements.Count; c++)
                    {
                        XmlNode imported = document.ImportNode(
                            replacements[c],
                            true);
                        parent.InsertBefore(imported, alternate);
                    }
                }

                parent.RemoveChild(alternate);
            }
        }

        private static void CollectAlternateContent(
            XmlNode node,
            List<XmlNode> result)
        {
            if (node == null)
                return;

            if (node.LocalName == "AlternateContent")
                result.Add(node);

            for (int i = 0; i < node.ChildNodes.Count; i++)
                CollectAlternateContent(node.ChildNodes[i], result);
        }

        private static XmlNode DirectChild(XmlNode node, string localName)
        {
            if (node == null)
                return null;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                XmlNode child = node.ChildNodes[i];
                if (child.LocalName == localName)
                    return child;
            }

            return null;
        }

        private static bool IsPresentationXml(string fullName)
        {
            if (string.IsNullOrEmpty(fullName))
                return false;

            string name = fullName.Replace('\\', '/');

            return name.StartsWith(
                       "ppt/",
                       StringComparison.OrdinalIgnoreCase) &&
                   name.EndsWith(
                       ".xml",
                       StringComparison.OrdinalIgnoreCase) &&
                   name.IndexOf(
                       "/_rels/",
                       StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static string BuildCacheKey(string path)
        {
            FileInfo info = new FileInfo(path);
            string value =
                info.FullName.ToLowerInvariant() + "|" +
                info.Length.ToString() + "|" +
                info.LastWriteTimeUtc.Ticks.ToString() +
                "|alternate-content-v1";

            byte[] data = Encoding.UTF8.GetBytes(value);

            using (SHA1 sha = SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(data);
                StringBuilder builder = new StringBuilder(hash.Length * 2);

                for (int i = 0; i < hash.Length; i++)
                    builder.Append(hash[i].ToString("x2"));

                return builder.ToString();
            }
        }

        private static void DeleteIfExists(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
            }
        }
    }
}
