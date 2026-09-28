using System;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal static class DocxReader
    {
        public static TextDocument Read(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("DOCX file was not found.", path);

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                XmlDocument documentXml =
                    OpcPackageUtility.ReadXmlPart(archive, "word/document.xml");

                if (documentXml == null)
                    throw new InvalidDataException("word/document.xml is missing or invalid.");

                TextDocument document = new TextDocument();
                document.Title = ReadCoreTitle(archive);

                XmlNode body = FindFirst(documentXml.DocumentElement, "body");
                if (body == null)
                    return document;

                for (int i = 0; i < body.ChildNodes.Count; i++)
                {
                    XmlNode node = body.ChildNodes[i];
                    if (node.LocalName != "p")
                        continue;

                    document.Paragraphs.Add(ReadParagraph(node));
                }

                return document;
            }
        }

        private static DocumentParagraph ReadParagraph(XmlNode paragraphNode)
        {
            DocumentParagraph paragraph = new DocumentParagraph();
            XmlNode properties = FindDirectChild(paragraphNode, "pPr");
            XmlNode justification = FindFirst(properties, "jc");
            paragraph.Alignment = ParseAlignment(
                GetAttribute(justification, "val"));

            for (int i = 0; i < paragraphNode.ChildNodes.Count; i++)
            {
                XmlNode runNode = paragraphNode.ChildNodes[i];
                if (runNode.LocalName != "r")
                    continue;

                DocumentTextRun run = ReadRun(runNode);
                if (run != null)
                    paragraph.Runs.Add(run);
            }

            return paragraph;
        }

        private static DocumentTextRun ReadRun(XmlNode runNode)
        {
            DocumentTextRun run = new DocumentTextRun();
            XmlNode properties = FindDirectChild(runNode, "rPr");

            if (properties != null)
            {
                XmlNode fonts = FindFirst(properties, "rFonts");
                string family = GetAttribute(fonts, "ascii");
                if (string.IsNullOrEmpty(family))
                    family = GetAttribute(fonts, "hAnsi");
                if (string.IsNullOrEmpty(family))
                    family = GetAttribute(fonts, "eastAsia");
                if (!string.IsNullOrEmpty(family))
                    run.FontFamily = family;

                run.Bold = FindDirectChild(properties, "b") != null;
                run.Italic = FindDirectChild(properties, "i") != null;

                XmlNode underline = FindDirectChild(properties, "u");
                if (underline != null)
                {
                    string value = GetAttribute(underline, "val");
                    run.Underline =
                        string.IsNullOrEmpty(value) ||
                        !string.Equals(value, "none", StringComparison.OrdinalIgnoreCase);
                }

                XmlNode color = FindDirectChild(properties, "color");
                string colorValue = GetAttribute(color, "val");
                if (IsHexColor(colorValue))
                    run.ColorHex = colorValue.ToUpperInvariant();

                XmlNode size = FindDirectChild(properties, "sz");
                int halfPoints;
                if (int.TryParse(GetAttribute(size, "val"), out halfPoints) &&
                    halfPoints > 0)
                {
                    run.FontSizePoints = halfPoints / 2.0f;
                }
            }

            string text = string.Empty;

            for (int i = 0; i < runNode.ChildNodes.Count; i++)
            {
                XmlNode child = runNode.ChildNodes[i];

                if (child.LocalName == "t")
                    text += child.InnerText;
                else if (child.LocalName == "tab")
                    text += "\t";
                else if (child.LocalName == "br" || child.LocalName == "cr")
                    text += "\n";
            }

            run.Text = text;
            return run;
        }

        private static string ReadCoreTitle(ZipArchive archive)
        {
            XmlDocument core =
                OpcPackageUtility.ReadXmlPart(archive, "docProps/core.xml");

            if (core == null)
                return "Document";

            XmlNode title = FindFirst(core.DocumentElement, "title");
            return title == null || string.IsNullOrEmpty(title.InnerText)
                ? "Document"
                : title.InnerText;
        }

        private static DocumentParagraphAlignment ParseAlignment(string value)
        {
            if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase))
                return DocumentParagraphAlignment.Center;
            if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "end", StringComparison.OrdinalIgnoreCase))
                return DocumentParagraphAlignment.Right;
            if (string.Equals(value, "both", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "distribute", StringComparison.OrdinalIgnoreCase))
                return DocumentParagraphAlignment.Justify;
            return DocumentParagraphAlignment.Left;
        }

        private static XmlNode FindDirectChild(XmlNode node, string localName)
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

        private static XmlNode FindFirst(XmlNode node, string localName)
        {
            if (node == null)
                return null;

            if (node.LocalName == localName)
                return node;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                XmlNode found = FindFirst(node.ChildNodes[i], localName);
                if (found != null)
                    return found;
            }

            return null;
        }

        private static string GetAttribute(XmlNode node, string localName)
        {
            if (node == null || node.Attributes == null)
                return string.Empty;

            for (int i = 0; i < node.Attributes.Count; i++)
            {
                XmlAttribute attribute = node.Attributes[i];
                if (attribute.LocalName == localName)
                    return attribute.Value ?? string.Empty;
            }

            return string.Empty;
        }

        private static bool IsHexColor(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length != 6)
                return false;

            for (int i = 0; i < value.Length; i++)
            {
                char ch = value[i];
                bool valid =
                    (ch >= '0' && ch <= '9') ||
                    (ch >= 'A' && ch <= 'F') ||
                    (ch >= 'a' && ch <= 'f');
                if (!valid)
                    return false;
            }

            return true;
        }
    }
}
