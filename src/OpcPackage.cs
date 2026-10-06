using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal sealed class OpcRelationship
    {
        public string Id { get; set; }
        public string Type { get; set; }
        public string Target { get; set; }
        public string TargetMode { get; set; }
        public string SourcePart { get; set; }
        public string ResolvedPart { get; set; }

        public bool IsExternal
        {
            get
            {
                return string.Equals(
                    TargetMode,
                    "External",
                    StringComparison.OrdinalIgnoreCase);
            }
        }

        public OpcRelationship()
        {
            Id = string.Empty;
            Type = string.Empty;
            Target = string.Empty;
            TargetMode = string.Empty;
            SourcePart = string.Empty;
            ResolvedPart = string.Empty;
        }
    }

    internal sealed class OpcContentTypeMap
    {
        private readonly Dictionary<string, string> defaults =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, string> overrides =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public void AddDefault(string extension, string contentType)
        {
            if (string.IsNullOrEmpty(extension) || string.IsNullOrEmpty(contentType))
                return;

            extension = extension.TrimStart('.');
            defaults[extension] = contentType;
        }

        public void AddOverride(string partName, string contentType)
        {
            if (string.IsNullOrEmpty(partName) || string.IsNullOrEmpty(contentType))
                return;

            overrides[OpcPackageUtility.NormalizePartName(partName)] = contentType;
        }

        public string GetContentType(string partName)
        {
            string normalized = OpcPackageUtility.NormalizePartName(partName);
            string contentType;

            if (overrides.TryGetValue(normalized, out contentType))
                return contentType;

            string extension = Path.GetExtension(normalized);
            if (!string.IsNullOrEmpty(extension))
            {
                extension = extension.TrimStart('.');

                if (defaults.TryGetValue(extension, out contentType))
                    return contentType;
            }

            return null;
        }
    }

    internal static class OpcPackageUtility
    {
        private const string RelationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        private const string ContentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";

        public static string NormalizePartName(string partName)
        {
            if (string.IsNullOrEmpty(partName))
                return string.Empty;

            string value = partName.Replace('\\', '/');

            while (value.StartsWith("/", StringComparison.Ordinal))
                value = value.Substring(1);

            string[] pieces = value.Split('/');
            List<string> normalized = new List<string>();

            for (int i = 0; i < pieces.Length; i++)
            {
                string piece = pieces[i];

                if (string.IsNullOrEmpty(piece) || piece == ".")
                    continue;

                if (piece == "..")
                {
                    if (normalized.Count > 0)
                        normalized.RemoveAt(normalized.Count - 1);

                    continue;
                }

                normalized.Add(piece);
            }

            return string.Join("/", normalized.ToArray());
        }

        public static string GetRelationshipPartName(string sourcePart)
        {
            sourcePart = NormalizePartName(sourcePart);

            if (string.IsNullOrEmpty(sourcePart))
                return "_rels/.rels";

            int slash = sourcePart.LastIndexOf('/');
            string directory = slash >= 0
                ? sourcePart.Substring(0, slash + 1)
                : string.Empty;

            string fileName = slash >= 0
                ? sourcePart.Substring(slash + 1)
                : sourcePart;

            return directory + "_rels/" + fileName + ".rels";
        }

        public static string ResolveTargetPart(string sourcePart, string target)
        {
            if (string.IsNullOrEmpty(target))
                return string.Empty;

            string normalizedTarget = target.Replace('\\', '/');

            if (normalizedTarget.StartsWith("/", StringComparison.Ordinal))
                return NormalizePartName(normalizedTarget);

            sourcePart = NormalizePartName(sourcePart);

            int slash = sourcePart.LastIndexOf('/');
            string directory = slash >= 0
                ? sourcePart.Substring(0, slash + 1)
                : string.Empty;

            return NormalizePartName(directory + normalizedTarget);
        }

        public static XmlDocument ReadXmlPart(ZipArchive archive, string partName)
        {
            if (archive == null)
                throw new ArgumentNullException("archive");

            string normalized = NormalizePartName(partName);
            ZipArchiveEntry entry = archive.GetEntry(normalized);

            if (entry == null)
                return null;

            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = true;

            using (Stream stream = entry.Open())
                document.Load(stream);

            return document;
        }

        public static List<OpcRelationship> ReadRelationships(
            ZipArchive archive,
            string sourcePart)
        {
            List<OpcRelationship> result = new List<OpcRelationship>();
            string relationshipPart = GetRelationshipPartName(sourcePart);
            XmlDocument document = ReadXmlPart(archive, relationshipPart);

            if (document == null)
                return result;

            XmlNode root = document.DocumentElement;
            if (root == null)
                return result;

            for (int i = 0; i < root.ChildNodes.Count; i++)
            {
                XmlNode node = root.ChildNodes[i];

                if (node.LocalName != "Relationship")
                    continue;

                OpcRelationship relationship = new OpcRelationship();
                relationship.Id = GetAttribute(node, "Id");
                relationship.Type = GetAttribute(node, "Type");
                relationship.Target = GetAttribute(node, "Target");
                relationship.TargetMode = GetAttribute(node, "TargetMode");
                relationship.SourcePart = NormalizePartName(sourcePart);

                if (!relationship.IsExternal)
                {
                    relationship.ResolvedPart = ResolveTargetPart(
                        sourcePart,
                        relationship.Target);
                }

                result.Add(relationship);
            }

            return result;
        }

        public static OpcContentTypeMap ReadContentTypes(ZipArchive archive)
        {
            OpcContentTypeMap result = new OpcContentTypeMap();
            XmlDocument document = ReadXmlPart(archive, "[Content_Types].xml");

            if (document == null || document.DocumentElement == null)
                return result;

            XmlNode root = document.DocumentElement;

            for (int i = 0; i < root.ChildNodes.Count; i++)
            {
                XmlNode node = root.ChildNodes[i];

                if (node.LocalName == "Default")
                {
                    result.AddDefault(
                        GetAttribute(node, "Extension"),
                        GetAttribute(node, "ContentType"));
                }
                else if (node.LocalName == "Override")
                {
                    result.AddOverride(
                        GetAttribute(node, "PartName"),
                        GetAttribute(node, "ContentType"));
                }
            }

            return result;
        }

        public static void WriteXmlPart(
            ZipArchive archive,
            string partName,
            XmlDocument document)
        {
            if (archive == null)
                throw new ArgumentNullException("archive");

            if (document == null)
                throw new ArgumentNullException("document");

            string normalized = NormalizePartName(partName);
            DeleteEntryIfPresent(archive, normalized);

            ZipArchiveEntry entry = archive.CreateEntry(
                normalized,
                CompressionLevel.Optimal);

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;

            using (Stream stream = entry.Open())
            using (XmlWriter writer = XmlWriter.Create(stream, settings))
                document.Save(writer);
        }

        public static XmlDocument CreateRelationshipsDocument(
            IList<OpcRelationship> relationships)
        {
            XmlDocument document = new XmlDocument();
            XmlElement root = document.CreateElement(
                "Relationships",
                RelationshipsNamespace);
            document.AppendChild(root);

            if (relationships == null)
                return document;

            for (int i = 0; i < relationships.Count; i++)
            {
                OpcRelationship relationship = relationships[i];
                if (relationship == null)
                    continue;

                XmlElement element = document.CreateElement(
                    "Relationship",
                    RelationshipsNamespace);

                element.SetAttribute("Id", relationship.Id ?? string.Empty);
                element.SetAttribute("Type", relationship.Type ?? string.Empty);
                element.SetAttribute("Target", relationship.Target ?? string.Empty);

                if (!string.IsNullOrEmpty(relationship.TargetMode))
                    element.SetAttribute("TargetMode", relationship.TargetMode);

                root.AppendChild(element);
            }

            return document;
        }

        public static XmlDocument CreateContentTypesDocument(
            IDictionary<string, string> defaults,
            IDictionary<string, string> overrides)
        {
            XmlDocument document = new XmlDocument();
            XmlElement root = document.CreateElement(
                "Types",
                ContentTypesNamespace);
            document.AppendChild(root);

            if (defaults != null)
            {
                foreach (KeyValuePair<string, string> pair in defaults)
                {
                    XmlElement element = document.CreateElement(
                        "Default",
                        ContentTypesNamespace);
                    element.SetAttribute("Extension", pair.Key.TrimStart('.'));
                    element.SetAttribute("ContentType", pair.Value);
                    root.AppendChild(element);
                }
            }

            if (overrides != null)
            {
                foreach (KeyValuePair<string, string> pair in overrides)
                {
                    XmlElement element = document.CreateElement(
                        "Override",
                        ContentTypesNamespace);
                    element.SetAttribute(
                        "PartName",
                        "/" + NormalizePartName(pair.Key));
                    element.SetAttribute("ContentType", pair.Value);
                    root.AppendChild(element);
                }
            }

            return document;
        }

        private static void DeleteEntryIfPresent(
            ZipArchive archive,
            string partName)
        {
            if (archive == null)
                return;

            // .NET Framework ZipArchive throws NotSupportedException when
            // GetEntry/Entries is queried while the archive is in Create
            // mode. Fresh package writers use Create mode and do not have an
            // existing entry to remove, so skip lookup entirely.
            if (archive.Mode == ZipArchiveMode.Create)
                return;

            ZipArchiveEntry existing =
                archive.GetEntry(partName);

            if (existing != null)
                existing.Delete();
        }

        private static string GetAttribute(XmlNode node, string name)
        {
            if (node == null || node.Attributes == null)
                return string.Empty;

            XmlAttribute attribute = node.Attributes[name];
            return attribute == null ? string.Empty : attribute.Value;
        }
    }
}
