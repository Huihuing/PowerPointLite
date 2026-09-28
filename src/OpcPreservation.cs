using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal sealed class OpcPreservedPartInfo
    {
        public string Name { get; set; }
        public long Length { get; set; }

        public OpcPreservedPartInfo()
        {
            Name = string.Empty;
        }
    }

    internal sealed class OpcPackagePreservationSnapshot
    {
        private readonly List<OpcPreservedPartInfo> parts =
            new List<OpcPreservedPartInfo>();

        public string SourcePath { get; private set; }

        public IList<OpcPreservedPartInfo> Parts
        {
            get { return parts.AsReadOnly(); }
        }

        private OpcPackagePreservationSnapshot(string sourcePath)
        {
            SourcePath = sourcePath;
        }

        public static OpcPackagePreservationSnapshot Capture(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("OPC source package was not found.", path);

            string fullPath = Path.GetFullPath(path);
            OpcPackagePreservationSnapshot snapshot =
                new OpcPackagePreservationSnapshot(fullPath);

            using (ZipArchive archive = ZipFile.OpenRead(fullPath))
            {
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    ZipArchiveEntry entry = archive.Entries[i];
                    if (entry == null || string.IsNullOrEmpty(entry.FullName))
                        continue;

                    OpcPreservedPartInfo info = new OpcPreservedPartInfo();
                    info.Name = OpcPackageUtility.NormalizePartName(entry.FullName);
                    info.Length = entry.Length;
                    snapshot.parts.Add(info);
                }
            }

            return snapshot;
        }

        public bool ContainsPart(string partName)
        {
            string normalized = OpcPackageUtility.NormalizePartName(partName);

            for (int i = 0; i < parts.Count; i++)
            {
                if (string.Equals(
                        parts[i].Name,
                        normalized,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public XmlDocument ReadOriginalContentTypes()
        {
            using (ZipArchive archive = ZipFile.OpenRead(SourcePath))
                return OpcPackageUtility.ReadXmlPart(archive, "[Content_Types].xml");
        }

        public void CopyUnchangedParts(
            ZipArchive destination,
            ICollection<string> replacedParts)
        {
            if (destination == null)
                throw new ArgumentNullException("destination");

            HashSet<string> excluded = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);

            if (replacedParts != null)
            {
                foreach (string part in replacedParts)
                {
                    if (!string.IsNullOrEmpty(part))
                        excluded.Add(OpcPackageUtility.NormalizePartName(part));
                }
            }

            using (ZipArchive source = ZipFile.OpenRead(SourcePath))
            {
                for (int i = 0; i < source.Entries.Count; i++)
                {
                    ZipArchiveEntry input = source.Entries[i];
                    if (input == null || string.IsNullOrEmpty(input.FullName))
                        continue;

                    string normalized =
                        OpcPackageUtility.NormalizePartName(input.FullName);

                    if (excluded.Contains(normalized))
                        continue;

                    if (destination.GetEntry(normalized) != null)
                        continue;

                    ZipArchiveEntry output = destination.CreateEntry(
                        normalized,
                        CompressionLevel.Optimal);

                    using (Stream inputStream = input.Open())
                    using (Stream outputStream = output.Open())
                        inputStream.CopyTo(outputStream);
                }
            }
        }

        public List<string> FindPartsOutsideKnownPrefixes(
            string[] knownPrefixes)
        {
            List<string> result = new List<string>();

            for (int i = 0; i < parts.Count; i++)
            {
                string name = parts[i].Name;
                bool known = false;

                if (knownPrefixes != null)
                {
                    for (int p = 0; p < knownPrefixes.Length; p++)
                    {
                        string prefix = OpcPackageUtility.NormalizePartName(
                            knownPrefixes[p]);

                        if (!string.IsNullOrEmpty(prefix) &&
                            name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                        {
                            known = true;
                            break;
                        }
                    }
                }

                if (!known)
                    result.Add(name);
            }

            return result;
        }
    }
}
