using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace PptxViewer
{
    internal sealed class HwpDocumentInfo
    {
        public string Version { get; set; }
        public uint RawVersion { get; set; }
        public uint Flags { get; set; }
        public bool IsCompressed { get; set; }
        public bool IsPasswordEncrypted { get; set; }
        public bool IsDistributionDocument { get; set; }
        public bool IsDrmProtected { get; set; }
        public bool IsCertificateEncrypted { get; set; }
        public int SectionCount { get; set; }

        public HwpDocumentInfo()
        {
            Version = string.Empty;
        }

        public bool HasUnsupportedProtection
        {
            get
            {
                return IsPasswordEncrypted ||
                    IsDistributionDocument ||
                    IsDrmProtected ||
                    IsCertificateEncrypted;
            }
        }
    }

    internal sealed class HwpReadResult
    {
        public TextDocument Document { get; set; }
        public HwpDocumentInfo Info { get; set; }

        public HwpReadResult()
        {
            Document = new TextDocument();
            Info = new HwpDocumentInfo();
        }
    }

    internal static class HwpReader
    {
        private const string HeaderSignature = "HWP Document File";

        public static HwpReadResult Read(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("HWP file was not found.", path);

            using (CompoundFileReader compound = new CompoundFileReader(path))
            {
                if (!compound.StreamExists("FileHeader"))
                    throw new InvalidDataException("HWP FileHeader stream is missing.");

                byte[] fileHeader = compound.ReadStream("FileHeader");
                HwpDocumentInfo info = ParseFileHeader(fileHeader);

                if (info.HasUnsupportedProtection)
                {
                    throw new NotSupportedException(
                        "This HWP uses password/distribution/DRM/certificate protection. " +
                        "The read-only foundation does not bypass or decrypt protected documents.");
                }

                TextDocument document = TextDocument.CreateNew(
                    Path.GetFileNameWithoutExtension(path));

                List<string> sections = FindBodySections(compound.GetStreamPaths());
                info.SectionCount = sections.Count;

                for (int i = 0; i < sections.Count; i++)
                {
                    byte[] section = compound.ReadStream(sections[i]);
                    if (info.IsCompressed)
                        section = Inflate(section);

                    List<string> paragraphs =
                        HwpRecordParser.ExtractParagraphTexts(section);

                    for (int p = 0; p < paragraphs.Count; p++)
                        document.AddParagraph(paragraphs[p]);
                }

                if (document.Paragraphs.Count == 0 && compound.StreamExists("PrvText"))
                {
                    string preview = DecodePreviewText(compound.ReadStream("PrvText"));
                    AddPreviewParagraphs(document, preview);
                }

                if (document.Paragraphs.Count == 0)
                    document.AddParagraph(string.Empty);

                HwpReadResult result = new HwpReadResult();
                result.Document = document;
                result.Info = info;
                return result;
            }
        }

        public static HwpDocumentInfo ReadInfo(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("HWP file was not found.", path);

            using (CompoundFileReader compound = new CompoundFileReader(path))
            {
                return ParseFileHeader(compound.ReadStream("FileHeader"));
            }
        }

        internal static HwpDocumentInfo ParseFileHeader(byte[] data)
        {
            if (data == null || data.Length < 40)
                throw new InvalidDataException("HWP FileHeader is truncated.");

            string signature = Encoding.ASCII
                .GetString(data, 0, Math.Min(32, data.Length))
                .TrimEnd('\0', ' ');

            if (!signature.StartsWith(HeaderSignature, StringComparison.Ordinal))
                throw new InvalidDataException("HWP FileHeader signature is invalid.");

            uint version = ReadUInt32(data, 32);
            uint flags = ReadUInt32(data, 36);

            HwpDocumentInfo info = new HwpDocumentInfo();
            info.RawVersion = version;
            info.Version =
                data[35].ToString(CultureInfo.InvariantCulture) + "." +
                data[34].ToString(CultureInfo.InvariantCulture) + "." +
                data[33].ToString(CultureInfo.InvariantCulture) + "." +
                data[32].ToString(CultureInfo.InvariantCulture);
            info.Flags = flags;
            info.IsCompressed = (flags & 0x00000001U) != 0;
            info.IsPasswordEncrypted = (flags & 0x00000002U) != 0;
            info.IsDistributionDocument = (flags & 0x00000004U) != 0;
            info.IsDrmProtected = (flags & 0x00000010U) != 0;
            info.IsCertificateEncrypted = (flags & 0x00000100U) != 0;
            return info;
        }

        private static List<string> FindBodySections(IList<string> paths)
        {
            List<SectionPath> sections = new List<SectionPath>();

            if (paths != null)
            {
                for (int i = 0; i < paths.Count; i++)
                {
                    string path = (paths[i] ?? string.Empty).Replace('\\', '/');
                    if (!path.StartsWith("BodyText/Section", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string suffix = path.Substring("BodyText/Section".Length);
                    int index;
                    if (!int.TryParse(
                            suffix,
                            NumberStyles.Integer,
                            CultureInfo.InvariantCulture,
                            out index))
                    {
                        continue;
                    }

                    SectionPath item = new SectionPath();
                    item.Index = index;
                    item.Path = path;
                    sections.Add(item);
                }
            }

            sections.Sort(delegate(SectionPath left, SectionPath right)
            {
                return left.Index.CompareTo(right.Index);
            });

            List<string> result = new List<string>();
            for (int i = 0; i < sections.Count; i++)
                result.Add(sections[i].Path);
            return result;
        }

        private sealed class SectionPath
        {
            public int Index;
            public string Path;
        }

        private static byte[] Inflate(byte[] data)
        {
            if (data == null || data.Length == 0)
                return new byte[0];

            try
            {
                using (MemoryStream input = new MemoryStream(data, false))
                using (DeflateStream inflater =
                    new DeflateStream(input, CompressionMode.Decompress, false))
                using (MemoryStream output = new MemoryStream())
                {
                    byte[] buffer = new byte[8192];
                    int read;
                    while ((read = inflater.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        output.Write(buffer, 0, read);
                        if (output.Length > 256L * 1024L * 1024L)
                            throw new InvalidDataException("Expanded HWP section exceeds the safety limit.");
                    }
                    return output.ToArray();
                }
            }
            catch (InvalidDataException)
            {
                throw new InvalidDataException(
                    "A compressed HWP stream could not be decompressed. " +
                    "The document may use an unsupported/corrupt stream encoding.");
            }
        }

        private static string DecodePreviewText(byte[] data)
        {
            if (data == null || data.Length == 0)
                return string.Empty;

            try
            {
                return Encoding.Unicode.GetString(data).TrimEnd('\0');
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void AddPreviewParagraphs(TextDocument document, string preview)
        {
            string normalized = (preview ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');
            string[] lines = normalized.Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Length > 0)
                    document.AddParagraph(lines[i]);
            }
        }

        private static uint ReadUInt32(byte[] data, int offset)
        {
            if (data == null || offset < 0 || offset + 4 > data.Length)
                throw new InvalidDataException("HWP integer is truncated.");

            return (uint)(
                data[offset] |
                (data[offset + 1] << 8) |
                (data[offset + 2] << 16) |
                (data[offset + 3] << 24));
        }
    }

    internal static class HwpRecordParser
    {
        // HWP 5.x BodyText record tag: HWPTAG_BEGIN(0x10) + 51.
        internal const int ParagraphTextTag = 67;

        public static List<string> ExtractParagraphTexts(byte[] data)
        {
            List<string> result = new List<string>();
            if (data == null || data.Length == 0)
                return result;

            int offset = 0;
            int recordCount = 0;

            while (offset + 4 <= data.Length)
            {
                if (++recordCount > 2000000)
                    throw new InvalidDataException("HWP record count exceeds the safety limit.");

                uint header = ReadUInt32(data, offset);
                offset += 4;

                int tagId = (int)(header & 0x3FFU);
                int size = (int)((header >> 20) & 0xFFFU);

                if (size == 0xFFF)
                {
                    if (offset + 4 > data.Length)
                        throw new InvalidDataException("Extended HWP record size is truncated.");
                    uint extended = ReadUInt32(data, offset);
                    offset += 4;
                    if (extended > int.MaxValue)
                        throw new InvalidDataException("HWP record is too large.");
                    size = (int)extended;
                }

                if (size < 0 || offset + size > data.Length)
                    throw new InvalidDataException("HWP record payload is truncated.");

                if (tagId == ParagraphTextTag)
                {
                    string text = DecodeParagraphText(data, offset, size);
                    result.Add(text);
                }

                offset += size;
            }

            return result;
        }

        internal static string DecodeParagraphText(byte[] data, int offset, int size)
        {
            if (data == null || size <= 0 || offset < 0 || offset + size > data.Length)
                return string.Empty;

            int evenSize = size - (size % 2);
            string raw = Encoding.Unicode.GetString(data, offset, evenSize);
            StringBuilder text = new StringBuilder();

            for (int i = 0; i < raw.Length; i++)
            {
                int code = raw[i];

                if (code >= 0x20)
                {
                    if (code != 0xFFFE && code != 0xFFFF)
                        text.Append(raw[i]);
                    continue;
                }

                if (code == 0x000A)
                {
                    text.Append('\n');
                    continue;
                }

                if (code == 0x0009)
                {
                    text.Append('\t');
                    i = SkipExtendedControl(raw, i);
                    continue;
                }

                if (code == 0x001E)
                {
                    text.Append('-');
                    continue;
                }

                if (code == 0x001F)
                {
                    text.Append(' ');
                    continue;
                }

                if (IsExtendedControl(code))
                {
                    i = SkipExtendedControl(raw, i);
                    continue;
                }

                // Other control codes are structural. Paragraph boundaries are
                // represented by PARA_TEXT records, so they are not emitted.
            }

            return text.ToString().TrimEnd('\0');
        }

        private static bool IsExtendedControl(int code)
        {
            return code == 0x0002 ||
                code == 0x0003 ||
                code == 0x0008 ||
                code == 0x000B ||
                code == 0x000F ||
                code == 0x0010 ||
                code == 0x0011 ||
                code == 0x0012 ||
                code == 0x0014 ||
                code == 0x0015;
        }

        private static int SkipExtendedControl(string raw, int index)
        {
            // Extended inline controls occupy eight UTF-16 code units in the
            // paragraph text stream. Keep this conservative and bounds-safe.
            int target = index + 7;
            return target < raw.Length ? target : raw.Length - 1;
        }

        private static uint ReadUInt32(byte[] data, int offset)
        {
            return (uint)(
                data[offset] |
                (data[offset + 1] << 8) |
                (data[offset + 2] << 16) |
                (data[offset + 3] << 24));
        }
    }
}
