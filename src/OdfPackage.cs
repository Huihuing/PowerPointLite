using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal enum OdfDocumentKind
    {
        Text,
        Spreadsheet,
        Presentation
    }

    internal static class OdfPackageUtility
    {
        public const string TextMimeType =
            "application/vnd.oasis.opendocument.text";
        public const string SpreadsheetMimeType =
            "application/vnd.oasis.opendocument.spreadsheet";
        public const string PresentationMimeType =
            "application/vnd.oasis.opendocument.presentation";

        public static string GetMimeType(OdfDocumentKind kind)
        {
            if (kind == OdfDocumentKind.Spreadsheet)
                return SpreadsheetMimeType;
            if (kind == OdfDocumentKind.Presentation)
                return PresentationMimeType;
            return TextMimeType;
        }

        public static bool HasExpectedMimeType(string path, OdfDocumentKind kind)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return false;

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                return string.Equals(
                    ReadTextEntry(archive, "mimetype").Trim(),
                    GetMimeType(kind),
                    StringComparison.Ordinal);
            }
        }

        public static XmlDocument LoadXml(ZipArchive archive, string part)
        {
            if (archive == null)
                throw new ArgumentNullException("archive");

            ZipArchiveEntry entry = archive.GetEntry(NormalizePart(part));
            if (entry == null)
                return null;

            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = true;
            using (Stream stream = entry.Open())
                document.Load(stream);
            return document;
        }

        public static string ReadTextEntry(ZipArchive archive, string part)
        {
            if (archive == null)
                throw new ArgumentNullException("archive");

            ZipArchiveEntry entry = archive.GetEntry(NormalizePart(part));
            if (entry == null)
                return string.Empty;

            using (Stream stream = entry.Open())
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
                return reader.ReadToEnd();
        }

        public static void WriteTextEntry(
            ZipArchive archive,
            string part,
            string content,
            CompressionLevel level)
        {
            if (archive == null)
                throw new ArgumentNullException("archive");

            ZipArchiveEntry entry = archive.CreateEntry(
                NormalizePart(part),
                level);

            using (Stream stream = entry.Open())
            using (StreamWriter writer =
                new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(content ?? string.Empty);
            }
        }

        public static string BuildManifest(
            OdfDocumentKind kind,
            IList<OdfManifestEntry> entries)
        {
            StringBuilder xml = new StringBuilder();
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            xml.Append("<manifest:manifest ");
            xml.Append("xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" ");
            xml.Append("manifest:version=\"1.3\">");
            xml.Append("<manifest:file-entry manifest:full-path=\"/\" manifest:version=\"1.3\" manifest:media-type=\"");
            xml.Append(EscapeXml(GetMimeType(kind)));
            xml.Append("\"/>");

            if (entries != null)
            {
                for (int i = 0; i < entries.Count; i++)
                {
                    OdfManifestEntry item = entries[i];
                    if (item == null || string.IsNullOrEmpty(item.Path))
                        continue;

                    xml.Append("<manifest:file-entry manifest:full-path=\"");
                    xml.Append(EscapeXml(item.Path));
                    xml.Append("\" manifest:media-type=\"");
                    xml.Append(EscapeXml(item.MediaType ?? string.Empty));
                    xml.Append("\"/>");
                }
            }

            xml.Append("</manifest:manifest>");
            return xml.ToString();
        }

        public static string BuildMetaXml(string title)
        {
            string safeTitle = EscapeXml(title ?? string.Empty);
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<office:document-meta " +
                "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
                "xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" office:version=\"1.3\">" +
                "<office:meta>" +
                "<meta:generator>PowerPointLite independent ODF writer</meta:generator>" +
                "<dc:title>" + safeTitle + "</dc:title>" +
                "<meta:creation-date>" + now + "</meta:creation-date>" +
                "<dc:date>" + now + "</dc:date>" +
                "</office:meta></office:document-meta>";
        }

        public static bool IsProjectGenerated(ZipArchive archive)
        {
            XmlDocument meta = LoadXml(archive, "meta.xml");
            if (meta == null || meta.DocumentElement == null)
                return false;

            XmlNodeList nodes = meta.GetElementsByTagName("*");
            for (int i = 0; i < nodes.Count; i++)
            {
                XmlNode node = nodes[i];
                if (node != null && node.LocalName == "generator" &&
                    (node.InnerText ?? string.Empty).IndexOf(
                        "PowerPointLite",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool ContainsOnlyKnownParts(
            ZipArchive archive,
            IList<string> allowedPrefixesOrParts)
        {
            if (archive == null)
                return false;

            for (int i = 0; i < archive.Entries.Count; i++)
            {
                string part = NormalizePart(archive.Entries[i].FullName);
                bool allowed = false;

                if (allowedPrefixesOrParts != null)
                {
                    for (int p = 0; p < allowedPrefixesOrParts.Count; p++)
                    {
                        string candidate = NormalizePart(allowedPrefixesOrParts[p]);
                        if (candidate.EndsWith("/", StringComparison.Ordinal))
                        {
                            if (part.StartsWith(candidate, StringComparison.OrdinalIgnoreCase))
                            {
                                allowed = true;
                                break;
                            }
                        }
                        else if (string.Equals(part, candidate, StringComparison.OrdinalIgnoreCase))
                        {
                            allowed = true;
                            break;
                        }
                    }
                }

                if (!allowed)
                    return false;
            }

            return true;
        }

        public static void ReplaceSafely(
            string stage,
            string destination,
            string backup)
        {
            if (!File.Exists(destination))
            {
                File.Move(stage, destination);
                return;
            }

            DeleteIfExists(backup);
            File.Move(destination, backup);

            try
            {
                File.Move(stage, destination);
                DeleteIfExists(backup);
            }
            catch
            {
                try
                {
                    DeleteIfExists(destination);
                    if (File.Exists(backup))
                        File.Move(backup, destination);
                }
                catch { }
                throw;
            }
        }

        public static void DeleteIfExists(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch { }
        }

        public static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }

        public static string GetAttributeByLocalName(XmlNode node, string localName)
        {
            if (node == null || node.Attributes == null)
                return string.Empty;

            for (int i = 0; i < node.Attributes.Count; i++)
            {
                XmlAttribute attribute = node.Attributes[i];
                if (attribute != null && attribute.LocalName == localName)
                    return attribute.Value;
            }

            return string.Empty;
        }

        private static string NormalizePart(string part)
        {
            return (part ?? string.Empty).Replace('\\', '/').TrimStart('/');
        }
    }

    internal sealed class OdfManifestEntry
    {
        public string Path { get; set; }
        public string MediaType { get; set; }

        public OdfManifestEntry()
        {
            Path = string.Empty;
            MediaType = string.Empty;
        }

        public OdfManifestEntry(string path, string mediaType)
        {
            Path = path ?? string.Empty;
            MediaType = mediaType ?? string.Empty;
        }
    }
}
