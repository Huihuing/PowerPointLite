using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Text;

namespace PptxViewer
{
    internal enum FontEmbeddingLevel
    {
        Unknown,
        Installable,
        Editable,
        PreviewAndPrint,
        Restricted
    }

    internal sealed class FontLicenseInfo
    {
        public string FontPath { get; set; }
        public ushort RawFsType { get; set; }
        public FontEmbeddingLevel EmbeddingLevel { get; set; }
        public bool CanEmbed { get; set; }
        public bool CanEmbedForEditing { get; set; }
        public bool CanPreviewAndPrint { get; set; }
        public bool CanSubset { get; set; }
        public bool BitmapEmbeddingOnly { get; set; }
        public bool MetadataAvailable { get; set; }
        public bool RequiresLicenseTextReview { get; set; }
        public string Note { get; set; }

        public FontLicenseInfo()
        {
            FontPath = string.Empty;
            EmbeddingLevel = FontEmbeddingLevel.Unknown;
            CanEmbed = false;
            CanEmbedForEditing = false;
            CanPreviewAndPrint = false;
            CanSubset = false;
            BitmapEmbeddingOnly = false;
            MetadataAvailable = false;
            RequiresLicenseTextReview = true;
            Note = "Font metadata is advisory. The actual font license text takes priority.";
        }
    }

    internal static class SystemFontCatalog
    {
        public static List<string> GetInstalledFamilyNames()
        {
            List<string> names = new List<string>();

            using (InstalledFontCollection fonts = new InstalledFontCollection())
            {
                for (int i = 0; i < fonts.Families.Length; i++)
                {
                    string name = fonts.Families[i].Name;

                    if (!string.IsNullOrEmpty(name) &&
                        !names.Contains(name))
                    {
                        names.Add(name);
                    }
                }
            }

            names.Sort(StringComparer.CurrentCultureIgnoreCase);
            return names;
        }
    }

    internal static class OpenTypeFontLicenseReader
    {
        private const ushort RestrictedLicenseEmbedding = 0x0002;
        private const ushort PreviewAndPrintEmbedding = 0x0004;
        private const ushort EditableEmbedding = 0x0008;
        private const ushort NoSubsetting = 0x0100;
        private const ushort BitmapEmbeddingOnly = 0x0200;

        public static FontLicenseInfo Read(string fontPath)
        {
            FontLicenseInfo info = new FontLicenseInfo();
            info.FontPath = fontPath ?? string.Empty;

            if (string.IsNullOrEmpty(fontPath) || !File.Exists(fontPath))
            {
                info.Note = "Font file was not found. Embedding is denied until the file and license can be inspected.";
                return info;
            }

            try
            {
                using (FileStream stream = new FileStream(fontPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (BinaryReader reader = new BinaryReader(stream, Encoding.ASCII))
                {
                    long sfntOffset = ResolveSfntOffset(reader);

                    if (sfntOffset < 0)
                    {
                        info.Note = "Unsupported or invalid OpenType/TrueType container.";
                        return info;
                    }

                    long os2Offset;
                    long os2Length;

                    if (!TryFindTable(reader, sfntOffset, "OS/2", out os2Offset, out os2Length))
                    {
                        info.Note = "OS/2 table is missing. Embedding permission cannot be inferred from fsType.";
                        return info;
                    }

                    if (os2Length < 10 || os2Offset < 0 || os2Offset + 10 > stream.Length)
                    {
                        info.Note = "OS/2 table is truncated or invalid.";
                        return info;
                    }

                    stream.Position = os2Offset + 8;
                    ushort fsType = ReadUInt16BigEndian(reader);
                    ApplyFsType(info, fsType);
                }
            }
            catch (Exception ex)
            {
                info.MetadataAvailable = false;
                info.CanEmbed = false;
                info.CanEmbedForEditing = false;
                info.CanPreviewAndPrint = false;
                info.CanSubset = false;
                info.EmbeddingLevel = FontEmbeddingLevel.Unknown;
                info.Note = "Font metadata could not be read: " + ex.Message;
            }

            return info;
        }

        private static long ResolveSfntOffset(BinaryReader reader)
        {
            Stream stream = reader.BaseStream;

            if (stream.Length < 12)
                return -1;

            stream.Position = 0;
            uint signature = ReadUInt32BigEndian(reader);

            // TrueType, OpenType/CFF, Apple TrueType and Type 1 sfnt signatures.
            if (signature == 0x00010000 ||
                signature == MakeTag("OTTO") ||
                signature == MakeTag("true") ||
                signature == MakeTag("typ1"))
            {
                return 0;
            }

            // TrueType Collection. Inspect the first font face only for now.
            if (signature == MakeTag("ttcf"))
            {
                if (stream.Length < 16)
                    return -1;

                ReadUInt32BigEndian(reader); // TTC version
                uint fontCount = ReadUInt32BigEndian(reader);

                if (fontCount == 0 || stream.Position + 4 > stream.Length)
                    return -1;

                uint firstOffset = ReadUInt32BigEndian(reader);

                if (firstOffset >= stream.Length)
                    return -1;

                return firstOffset;
            }

            return -1;
        }

        private static bool TryFindTable(
            BinaryReader reader,
            long sfntOffset,
            string tableTag,
            out long tableOffset,
            out long tableLength)
        {
            tableOffset = -1;
            tableLength = 0;

            Stream stream = reader.BaseStream;

            if (sfntOffset < 0 || sfntOffset + 12 > stream.Length)
                return false;

            stream.Position = sfntOffset + 4;
            ushort tableCount = ReadUInt16BigEndian(reader);

            long directoryStart = sfntOffset + 12;
            long directoryLength = tableCount * 16L;

            if (directoryStart + directoryLength > stream.Length)
                return false;

            uint wantedTag = MakeTag(tableTag);
            stream.Position = directoryStart;

            for (int i = 0; i < tableCount; i++)
            {
                uint tag = ReadUInt32BigEndian(reader);
                ReadUInt32BigEndian(reader); // checksum
                uint offset = ReadUInt32BigEndian(reader);
                uint length = ReadUInt32BigEndian(reader);

                if (tag == wantedTag)
                {
                    if (offset > stream.Length ||
                        length > stream.Length ||
                        (long)offset + (long)length > stream.Length)
                    {
                        return false;
                    }

                    tableOffset = offset;
                    tableLength = length;
                    return true;
                }
            }

            return false;
        }

        private static void ApplyFsType(FontLicenseInfo info, ushort fsType)
        {
            info.RawFsType = fsType;
            info.MetadataAvailable = true;
            info.RequiresLicenseTextReview = true;
            info.BitmapEmbeddingOnly = (fsType & BitmapEmbeddingOnly) != 0;
            info.CanSubset = (fsType & NoSubsetting) == 0;

            bool restricted = (fsType & RestrictedLicenseEmbedding) != 0;
            bool previewPrint = (fsType & PreviewAndPrintEmbedding) != 0;
            bool editable = (fsType & EditableEmbedding) != 0;

            if (restricted)
            {
                info.EmbeddingLevel = FontEmbeddingLevel.Restricted;
                info.CanEmbed = false;
                info.CanEmbedForEditing = false;
                info.CanPreviewAndPrint = false;
                info.Note = "fsType marks this font as Restricted License Embedding. Do not embed unless the actual license grants a separate permission.";
                return;
            }

            if (info.BitmapEmbeddingOnly)
            {
                info.EmbeddingLevel = FontEmbeddingLevel.Unknown;
                info.CanEmbed = false;
                info.CanEmbedForEditing = false;
                info.CanPreviewAndPrint = false;
                info.Note = "fsType allows bitmap-only embedding. Outline font embedding is disabled by this project.";
                return;
            }

            if (editable)
            {
                info.EmbeddingLevel = FontEmbeddingLevel.Editable;
                info.CanEmbed = true;
                info.CanEmbedForEditing = true;
                info.CanPreviewAndPrint = true;
                info.Note = "fsType indicates Editable Embedding. Verify the actual font license before redistribution or document embedding.";
                return;
            }

            if (previewPrint)
            {
                info.EmbeddingLevel = FontEmbeddingLevel.PreviewAndPrint;
                info.CanEmbed = true;
                info.CanEmbedForEditing = false;
                info.CanPreviewAndPrint = true;
                info.Note = "fsType indicates Preview & Print Embedding only. Editable-document embedding is disabled.";
                return;
            }

            // Per the OpenType specification, no embedding restriction bits means
            // Installable Embedding. The external license text still takes priority.
            info.EmbeddingLevel = FontEmbeddingLevel.Installable;
            info.CanEmbed = true;
            info.CanEmbedForEditing = true;
            info.CanPreviewAndPrint = true;
            info.Note = "fsType contains no embedding restriction bits (Installable Embedding). Verify the actual license text before bundling or redistribution.";
        }

        private static ushort ReadUInt16BigEndian(BinaryReader reader)
        {
            byte high = reader.ReadByte();
            byte low = reader.ReadByte();
            return (ushort)((high << 8) | low);
        }

        private static uint ReadUInt32BigEndian(BinaryReader reader)
        {
            uint b1 = reader.ReadByte();
            uint b2 = reader.ReadByte();
            uint b3 = reader.ReadByte();
            uint b4 = reader.ReadByte();

            return (b1 << 24) |
                   (b2 << 16) |
                   (b3 << 8) |
                   b4;
        }

        private static uint MakeTag(string value)
        {
            if (value == null || value.Length != 4)
                throw new ArgumentException("OpenType tag must contain exactly four characters.", "value");

            return ((uint)(byte)value[0] << 24) |
                   ((uint)(byte)value[1] << 16) |
                   ((uint)(byte)value[2] << 8) |
                   (uint)(byte)value[3];
        }
    }
}
