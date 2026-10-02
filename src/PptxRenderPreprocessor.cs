using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
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

            string tempRoot = Path.Combine(
                Path.GetTempPath(),
                "PowerPointLite",
                "RenderPrep");
            Directory.CreateDirectory(tempRoot);

            string destination = Path.Combine(
                tempRoot,
                Guid.NewGuid().ToString("N") + extension);
            string stage = destination + ".writing";

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

                File.Move(stage, destination);
                return destination;
            }
            catch
            {
                DeleteIfExists(stage);
                DeleteIfExists(destination);
                throw;
            }
        }

        public static void DeletePreparedFile(string path, string originalPath)
        {
            if (string.IsNullOrEmpty(path) ||
                string.Equals(path, originalPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            DeleteIfExists(path);
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

                // Prefer the compatibility fallback. PowerPoint commonly stores
                // a raster or broadly supported DrawingML representation here.
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
