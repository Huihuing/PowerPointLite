using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PptxViewer
{
    internal static class HwpDiagnostics
    {
        public static void ValidateParserFoundation()
        {
            ValidateFileHeaderParser();
            ValidateRecordParser();
            ValidateControlFiltering();
        }

        private static void ValidateFileHeaderParser()
        {
            byte[] header = new byte[256];
            byte[] signature = Encoding.ASCII.GetBytes("HWP Document File");
            Buffer.BlockCopy(signature, 0, header, 0, signature.Length);

            // Version 5.1.0.0 when interpreted in HWP byte order.
            header[32] = 0;
            header[33] = 0;
            header[34] = 1;
            header[35] = 5;

            // Compression flag only.
            header[36] = 1;

            HwpDocumentInfo info = HwpReader.ParseFileHeader(header);
            if (info == null ||
                info.Version != "5.1.0.0" ||
                !info.IsCompressed ||
                info.HasUnsupportedProtection)
            {
                throw new InvalidOperationException("HWP FileHeader parser self-test failed.");
            }
        }

        private static void ValidateRecordParser()
        {
            byte[] paragraph = Encoding.Unicode.GetBytes("Hello HWP 5.x");
            byte[] ignored = new byte[] { 1, 2, 3, 4 };

            using (MemoryStream stream = new MemoryStream())
            {
                WriteRecord(stream, 66, 0, ignored);
                WriteRecord(stream, HwpRecordParser.ParagraphTextTag, 1, paragraph);

                List<string> texts =
                    HwpRecordParser.ExtractParagraphTexts(stream.ToArray());

                if (texts.Count != 1 || texts[0] != "Hello HWP 5.x")
                    throw new InvalidOperationException("HWP PARA_TEXT record parser self-test failed.");
            }
        }

        private static void ValidateControlFiltering()
        {
            List<char> chars = new List<char>();
            chars.Add('A');
            chars.Add((char)0x000A);
            chars.Add('B');
            chars.Add((char)0x001E);
            chars.Add('C');
            chars.Add((char)0x001F);
            chars.Add('D');

            string value = new string(chars.ToArray());
            byte[] payload = Encoding.Unicode.GetBytes(value);
            string decoded = HwpRecordParser.DecodeParagraphText(payload, 0, payload.Length);

            if (decoded != "A\nB-C D")
                throw new InvalidOperationException("HWP inline control filtering self-test failed.");
        }

        private static void WriteRecord(
            Stream stream,
            int tagId,
            int level,
            byte[] payload)
        {
            int size = payload == null ? 0 : payload.Length;
            if (size >= 0xFFF)
                throw new InvalidOperationException("Synthetic diagnostic payload is too large.");

            uint header =
                ((uint)tagId & 0x3FFU) |
                (((uint)level & 0x3FFU) << 10) |
                (((uint)size & 0xFFFU) << 20);

            WriteUInt32(stream, header);
            if (size > 0)
                stream.Write(payload, 0, size);
        }

        private static void WriteUInt32(Stream stream, uint value)
        {
            stream.WriteByte((byte)(value & 0xFF));
            stream.WriteByte((byte)((value >> 8) & 0xFF));
            stream.WriteByte((byte)((value >> 16) & 0xFF));
            stream.WriteByte((byte)((value >> 24) & 0xFF));
        }
    }
}
