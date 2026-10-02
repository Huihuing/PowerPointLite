using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxRichTextPackage
    {
        private const string DrawingNamespace =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static void InjectRichText(
            PresentationDocument document,
            string packagePath)
        {
            if (document == null || string.IsNullOrEmpty(packagePath) || !File.Exists(packagePath))
                return;

            using (ZipArchive archive = ZipFile.Open(packagePath, ZipArchiveMode.Update))
            {
                for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
                {
                    PresentationSlide slide = document.Slides[slideIndex];
                    if (slide == null || !HasRichText(slide))
                        continue;

                    string partName = "ppt/slides/slide" + (slideIndex + 1).ToString() + ".xml";
                    ZipArchiveEntry entry = archive.GetEntry(partName);
                    if (entry == null)
                        continue;

                    XmlDocument xml = ReadEntry(entry);
                    List<XmlNode> textShapes = FindWriterTextShapes(xml);
                    int count = Math.Min(textShapes.Count, slide.TextBoxes.Count);

                    for (int i = 0; i < count; i++)
                    {
                        PresentationTextBox box = slide.TextBoxes[i];
                        if (box == null || !box.HasRichText)
                            continue;

                        XmlNode txBody = FindFirst(textShapes[i], "txBody");
                        if (txBody == null)
                            continue;

                        ReplaceParagraphs(xml, txBody, box);
                    }

                    ReplaceEntry(archive, partName, xml);
                }
            }
        }

        public static void ReadIntoDocument(
            string packagePath,
            PresentationDocument document)
        {
            if (document == null || string.IsNullOrEmpty(packagePath) || !File.Exists(packagePath))
                return;

            using (ZipArchive archive = ZipFile.OpenRead(packagePath))
            {
                for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
                {
                    PresentationSlide slide = document.Slides[slideIndex];
                    if (slide == null || slide.TextBoxes.Count == 0)
                        continue;

                    ZipArchiveEntry entry = archive.GetEntry(
                        "ppt/slides/slide" + (slideIndex + 1).ToString() + ".xml");
                    if (entry == null)
                        continue;

                    XmlDocument xml = ReadEntry(entry);
                    List<XmlNode> textShapes = FindWriterTextShapes(xml);
                    int count = Math.Min(textShapes.Count, slide.TextBoxes.Count);

                    for (int i = 0; i < count; i++)
                    {
                        PresentationTextBox box = slide.TextBoxes[i];
                        XmlNode txBody = FindFirst(textShapes[i], "txBody");
                        if (box == null || txBody == null)
                            continue;

                        List<PresentationTextParagraph> paragraphs =
                            ReadParagraphs(txBody, box);

                        if (paragraphs.Count > 0 && IsMeaningfullyRich(paragraphs, box))
                            box.SetRichParagraphs(paragraphs);
                    }
                }
            }
        }

        private static bool HasRichText(PresentationSlide slide)
        {
            for (int i = 0; i < slide.TextBoxes.Count; i++)
            {
                PresentationTextBox box = slide.TextBoxes[i];
                if (box != null && box.HasRichText)
                    return true;
            }
            return false;
        }

        private static List<XmlNode> FindWriterTextShapes(XmlDocument document)
        {
            List<XmlNode> result = new List<XmlNode>();
            XmlNode tree = FindFirst(document, "spTree");
            if (tree == null)
                return result;

            for (int i = 0; i < tree.ChildNodes.Count; i++)
            {
                XmlNode child = tree.ChildNodes[i];
                if (child.LocalName != "sp")
                    continue;

                XmlNode cNvSpPr = FindFirst(child, "cNvSpPr");
                string txBox = GetAttr(cNvSpPr, "txBox");
                if (txBox == "1" || string.Equals(txBox, "true", StringComparison.OrdinalIgnoreCase))
                    result.Add(child);
            }
            return result;
        }

        private static void ReplaceParagraphs(
            XmlDocument owner,
            XmlNode txBody,
            PresentationTextBox box)
        {
            for (int i = txBody.ChildNodes.Count - 1; i >= 0; i--)
            {
                if (txBody.ChildNodes[i].LocalName == "p")
                    txBody.RemoveChild(txBody.ChildNodes[i]);
            }

            XmlDocument fragment = new XmlDocument();
            fragment.LoadXml(
                "<root xmlns:a=\"" + DrawingNamespace + "\">" +
                BuildParagraphXml(box) +
                "</root>");

            while (fragment.DocumentElement.HasChildNodes)
            {
                XmlNode child = fragment.DocumentElement.FirstChild;
                fragment.DocumentElement.RemoveChild(child);
                txBody.AppendChild(owner.ImportNode(child, true));
            }
        }

        private static string BuildParagraphXml(PresentationTextBox box)
        {
            StringBuilder xml = new StringBuilder();

            for (int p = 0; p < box.RichParagraphs.Count; p++)
            {
                PresentationTextParagraph paragraph = box.RichParagraphs[p];
                if (paragraph == null)
                    continue;

                xml.Append("<a:p><a:pPr algn=\"");
                xml.Append(AlignmentValue(paragraph.Alignment));
                xml.Append("\" lvl=\"");
                xml.Append(Math.Max(0, Math.Min(8, paragraph.Level)).ToString());
                xml.Append("\">");

                if (!string.IsNullOrEmpty(paragraph.BulletText))
                {
                    xml.Append("<a:buChar char=\"");
                    xml.Append(EscapeXml(paragraph.BulletText));
                    xml.Append("\"/>");
                }

                AppendSpacing(xml, "spcBef", paragraph.SpaceBeforePoints);
                AppendSpacing(xml, "spcAft", paragraph.SpaceAfterPoints);
                xml.Append("</a:pPr>");

                for (int r = 0; r < paragraph.Runs.Count; r++)
                {
                    PresentationTextRun run = paragraph.Runs[r];
                    if (run == null)
                        continue;

                    if (run.Text == "\n")
                    {
                        xml.Append("<a:br/>");
                        continue;
                    }

                    xml.Append("<a:r>");
                    xml.Append(BuildRunProperties(run, box));
                    xml.Append("<a:t>");
                    xml.Append(EscapeXml(run.Text));
                    xml.Append("</a:t></a:r>");
                }

                xml.Append("<a:endParaRPr lang=\"ko-KR\" sz=\"");
                xml.Append(ToHundredths(box.FontSizePoints).ToString());
                xml.Append("\"/></a:p>");
            }

            if (xml.Length == 0)
                xml.Append("<a:p><a:endParaRPr lang=\"ko-KR\"/></a:p>");

            return xml.ToString();
        }

        private static string BuildRunProperties(
            PresentationTextRun run,
            PresentationTextBox box)
        {
            string family = string.IsNullOrEmpty(run.FontFamily)
                ? box.FontFamily
                : run.FontFamily;
            string color = NormalizeColor(
                string.IsNullOrEmpty(run.ColorHex)
                    ? box.ColorHex
                    : run.ColorHex);
            float size = run.FontSizePoints > 0f
                ? run.FontSizePoints
                : box.FontSizePoints;

            StringBuilder xml = new StringBuilder();
            xml.Append("<a:rPr lang=\"ko-KR\" sz=\"");
            xml.Append(ToHundredths(size).ToString());
            xml.Append("\"");
            if (run.Bold) xml.Append(" b=\"1\"");
            if (run.Italic) xml.Append(" i=\"1\"");
            if (run.Underline) xml.Append(" u=\"sng\"");
            if (run.BaselinePercent != 0)
            {
                xml.Append(" baseline=\"");
                xml.Append(Math.Max(-25000, Math.Min(25000, run.BaselinePercent * 1000)).ToString());
                xml.Append("\"");
            }
            xml.Append(">");
            xml.Append("<a:solidFill><a:srgbClr val=\"");
            xml.Append(color);
            xml.Append("\"/></a:solidFill>");
            xml.Append("<a:latin typeface=\"");
            xml.Append(EscapeXml(family));
            xml.Append("\"/><a:ea typeface=\"");
            xml.Append(EscapeXml(family));
            xml.Append("\"/></a:rPr>");
            return xml.ToString();
        }

        private static void AppendSpacing(
            StringBuilder xml,
            string element,
            float points)
        {
            if (points <= 0f)
                return;

            xml.Append("<a:");
            xml.Append(element);
            xml.Append("><a:spcPts val=\"");
            xml.Append(ToHundredths(points).ToString());
            xml.Append("\"/></a:");
            xml.Append(element);
            xml.Append(">");
        }

        private static List<PresentationTextParagraph> ReadParagraphs(
            XmlNode txBody,
            PresentationTextBox box)
        {
            List<PresentationTextParagraph> result =
                new List<PresentationTextParagraph>();

            for (int i = 0; i < txBody.ChildNodes.Count; i++)
            {
                XmlNode node = txBody.ChildNodes[i];
                if (node.LocalName != "p")
                    continue;

                PresentationTextParagraph paragraph = new PresentationTextParagraph();
                XmlNode pPr = DirectChild(node, "pPr");
                paragraph.Alignment = ParseAlignment(GetAttr(pPr, "algn"), box.Alignment);
                paragraph.Level = Math.Max(0, Math.Min(8, GetInt(GetAttr(pPr, "lvl"), 0)));

                XmlNode bullet = DirectChild(pPr, "buChar");
                paragraph.BulletText = GetAttr(bullet, "char") ?? string.Empty;
                paragraph.SpaceBeforePoints = ReadSpacing(pPr, "spcBef");
                paragraph.SpaceAfterPoints = ReadSpacing(pPr, "spcAft");

                for (int c = 0; c < node.ChildNodes.Count; c++)
                {
                    XmlNode child = node.ChildNodes[c];
                    if (child.LocalName == "r" || child.LocalName == "fld")
                    {
                        paragraph.Runs.Add(ReadRun(child, box));
                    }
                    else if (child.LocalName == "br")
                    {
                        PresentationTextRun lineBreak = CreateDefaultRun(box);
                        lineBreak.Text = "\n";
                        paragraph.Runs.Add(lineBreak);
                    }
                }

                if (paragraph.Runs.Count == 0)
                    paragraph.Runs.Add(CreateDefaultRun(box));

                result.Add(paragraph);
            }

            return result;
        }

        private static PresentationTextRun ReadRun(
            XmlNode runNode,
            PresentationTextBox box)
        {
            PresentationTextRun run = CreateDefaultRun(box);
            XmlNode rPr = DirectChild(runNode, "rPr");
            XmlNode text = DirectChild(runNode, "t");
            run.Text = text == null ? string.Empty : text.InnerText;

            int size = GetInt(GetAttr(rPr, "sz"), ToHundredths(box.FontSizePoints));
            run.FontSizePoints = Math.Max(1f, size / 100f);
            run.Bold = IsTrue(GetAttr(rPr, "b"));
            run.Italic = IsTrue(GetAttr(rPr, "i"));
            string underline = GetAttr(rPr, "u");
            run.Underline = !string.IsNullOrEmpty(underline) && underline != "none";
            run.BaselinePercent = GetInt(GetAttr(rPr, "baseline"), 0) / 1000;

            XmlNode latin = FindFirst(rPr, "latin");
            string family = GetAttr(latin, "typeface");
            if (!string.IsNullOrEmpty(family))
                run.FontFamily = family;

            XmlNode color = FindFirst(rPr, "srgbClr");
            string colorValue = GetAttr(color, "val");
            if (!string.IsNullOrEmpty(colorValue))
                run.ColorHex = colorValue;

            return run;
        }

        private static PresentationTextRun CreateDefaultRun(PresentationTextBox box)
        {
            PresentationTextRun run = new PresentationTextRun();
            run.FontFamily = box.FontFamily;
            run.FontSizePoints = box.FontSizePoints;
            run.Bold = box.Bold;
            run.Italic = box.Italic;
            run.ColorHex = box.ColorHex;
            return run;
        }

        private static bool IsMeaningfullyRich(
            IList<PresentationTextParagraph> paragraphs,
            PresentationTextBox box)
        {
            if (paragraphs.Count > 1)
                return true;

            for (int p = 0; p < paragraphs.Count; p++)
            {
                PresentationTextParagraph paragraph = paragraphs[p];
                if (paragraph == null)
                    continue;

                if (paragraph.Runs.Count > 1 ||
                    paragraph.Level != 0 ||
                    !string.IsNullOrEmpty(paragraph.BulletText) ||
                    paragraph.SpaceBeforePoints > 0f ||
                    paragraph.SpaceAfterPoints > 0f ||
                    paragraph.Alignment != box.Alignment)
                {
                    return true;
                }

                if (paragraph.Runs.Count == 1)
                {
                    PresentationTextRun run = paragraph.Runs[0];
                    if (run != null &&
                        (run.Underline || run.BaselinePercent != 0 ||
                         run.Bold != box.Bold || run.Italic != box.Italic ||
                         Math.Abs(run.FontSizePoints - box.FontSizePoints) > 0.01f ||
                         !string.Equals(run.FontFamily, box.FontFamily, StringComparison.OrdinalIgnoreCase) ||
                         !string.Equals(NormalizeColor(run.ColorHex), NormalizeColor(box.ColorHex), StringComparison.OrdinalIgnoreCase)))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static float ReadSpacing(XmlNode pPr, string name)
        {
            XmlNode spacing = DirectChild(pPr, name);
            XmlNode points = DirectChild(spacing, "spcPts");
            return Math.Max(0f, GetInt(GetAttr(points, "val"), 0) / 100f);
        }

        private static PresentationTextAlignment ParseAlignment(
            string value,
            PresentationTextAlignment fallback)
        {
            if (value == "ctr") return PresentationTextAlignment.Center;
            if (value == "r") return PresentationTextAlignment.Right;
            if (value == "l") return PresentationTextAlignment.Left;
            return fallback;
        }

        private static string AlignmentValue(PresentationTextAlignment value)
        {
            if (value == PresentationTextAlignment.Center) return "ctr";
            if (value == PresentationTextAlignment.Right) return "r";
            return "l";
        }

        private static int ToHundredths(float points)
        {
            return Math.Max(100, Math.Min(40000, (int)Math.Round(points * 100f)));
        }

        private static string NormalizeColor(string value)
        {
            string candidate = (value ?? string.Empty).Trim().TrimStart('#').ToUpperInvariant();
            if (candidate.Length != 6)
                return "20242A";
            int parsed;
            return int.TryParse(
                candidate,
                System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture,
                out parsed)
                ? candidate
                : "20242A";
        }

        private static bool IsTrue(string value)
        {
            return value == "1" ||
                string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        private static int GetInt(string value, int fallback)
        {
            int parsed;
            return !string.IsNullOrEmpty(value) && int.TryParse(value, out parsed)
                ? parsed
                : fallback;
        }

        private static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }

        private static XmlDocument ReadEntry(ZipArchiveEntry entry)
        {
            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = false;
            using (Stream input = entry.Open())
                document.Load(input);
            return document;
        }

        private static void ReplaceEntry(
            ZipArchive archive,
            string partName,
            XmlDocument document)
        {
            ZipArchiveEntry old = archive.GetEntry(partName);
            if (old != null)
                old.Delete();

            ZipArchiveEntry replacement = archive.CreateEntry(partName, CompressionLevel.Optimal);
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            using (Stream output = replacement.Open())
            using (XmlWriter writer = XmlWriter.Create(output, settings))
                document.Save(writer);
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

        private static XmlNode DirectChild(XmlNode node, string localName)
        {
            if (node == null)
                return null;
            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                if (node.ChildNodes[i].LocalName == localName)
                    return node.ChildNodes[i];
            }
            return null;
        }

        private static string GetAttr(XmlNode node, string name)
        {
            if (node == null || node.Attributes == null)
                return null;
            for (int i = 0; i < node.Attributes.Count; i++)
            {
                XmlAttribute attribute = node.Attributes[i];
                if (attribute.LocalName == name)
                    return attribute.Value;
            }
            return null;
        }
    }
}
