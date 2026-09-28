using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace PptxViewer
{
internal static partial class InternalPptxRenderer
    {
        private static void DrawShapeText(
            Graphics g,
            XmlNode shape,
            RectangleF rect,
            Dictionary<string, Color> theme,
            int slideNumber)
        {
            XmlNode txBody = DirectChild(shape, "txBody");
            if (txBody == null)
                return;

            XmlNode bodyPr = DirectChild(txBody, "bodyPr");
            float lIns = 6f;
            float rIns = 6f;
            float tIns = 4f;
            float bIns = 4f;

            if (bodyPr != null)
            {
                lIns = EmuToRenderPixels(GetLong(bodyPr, "lIns", 45720));
                rIns = EmuToRenderPixels(GetLong(bodyPr, "rIns", 45720));
                tIns = EmuToRenderPixels(GetLong(bodyPr, "tIns", 22860));
                bIns = EmuToRenderPixels(GetLong(bodyPr, "bIns", 22860));
            }

            RectangleF inner = new RectangleF(
                rect.X + lIns,
                rect.Y + tIns,
                Math.Max(1, rect.Width - lIns - rIns),
                Math.Max(1, rect.Height - tIns - bIns));

            float y = inner.Top;

            List<XmlNode> paragraphs = new List<XmlNode>();
            foreach (XmlNode child in txBody.ChildNodes)
            {
                if (child.LocalName == "p")
                    paragraphs.Add(child);
            }

            for (int pi = 0; pi < paragraphs.Count; pi++)
            {
                XmlNode paragraph = paragraphs[pi];

                StringBuilder text = new StringBuilder();
                foreach (XmlNode paragraphChild in paragraph.ChildNodes)
                {
                    if (paragraphChild.LocalName == "br")
                    {
                        text.AppendLine();
                        continue;
                    }

                    if (paragraphChild.LocalName == "r" || paragraphChild.LocalName == "fld")
                    {
                        foreach (XmlNode t in FindAll(paragraphChild, "t"))
                        {
                            if (paragraphChild.LocalName == "fld" &&
                                string.Equals(GetAttr(paragraphChild, "type"), "slidenum", StringComparison.OrdinalIgnoreCase))
                            {
                                text.Append(slideNumber.ToString());
                            }
                            else
                            {
                                text.Append(t.InnerText);
                            }
                        }
                    }
                }

                if (text.Length == 0)
                {
                    y += 8f;
                    continue;
                }

                XmlNode pPr = DirectChild(paragraph, "pPr");

                if (pPr != null)
                {
                    XmlNode buChar = DirectChild(pPr, "buChar");

                    if (buChar != null)
                    {
                        string bullet = GetAttr(buChar, "char");
                        if (string.IsNullOrEmpty(bullet))
                            bullet = "•";

                        text.Insert(0, bullet + " ");
                    }
                    else if (DirectChild(pPr, "buAutoNum") != null)
                    {
                        text.Insert(0, (pi + 1).ToString() + ". ");
                    }
                }

                XmlNode rPr = FindFirst(paragraph, "rPr");
                XmlNode endPr = DirectChild(paragraph, "endParaRPr");

                if (rPr == null)
                    rPr = endPr;

                float sizePt = 18f;
                string fontName = "Arial";
                Color color = Color.Black;
                FontStyle style = FontStyle.Regular;

                if (rPr != null)
                {
                    long sz = GetLong(rPr, "sz", 0);

                    if (sz > 0)
                        sizePt = (float)(sz / 100.0);

                    string bold = GetAttr(rPr, "b");
                    string italic = GetAttr(rPr, "i");
                    string underline = GetAttr(rPr, "u");

                    if (bold == "1" || string.Equals(bold, "true", StringComparison.OrdinalIgnoreCase))
                        style |= FontStyle.Bold;

                    if (italic == "1" || string.Equals(italic, "true", StringComparison.OrdinalIgnoreCase))
                        style |= FontStyle.Italic;

                    if (!string.IsNullOrEmpty(underline) && underline != "none")
                        style |= FontStyle.Underline;

                    XmlNode latin = ContainsCjk(text.ToString())
                        ? FindFirst(rPr, "ea")
                        : FindFirst(rPr, "latin");

                    if (latin == null)
                        latin = FindFirst(rPr, "latin");

                    if (latin != null)
                    {
                        string typeface = GetAttr(latin, "typeface");

                        if (!string.IsNullOrWhiteSpace(typeface))
                        {
                            if (typeface.StartsWith("+mj-ea") && ThemeFonts.ContainsKey("majorEA"))
                                fontName = ThemeFonts["majorEA"];
                            else if (typeface.StartsWith("+mn-ea") && ThemeFonts.ContainsKey("minorEA"))
                                fontName = ThemeFonts["minorEA"];
                            else if (typeface.StartsWith("+mj") && ThemeFonts.ContainsKey("major"))
                                fontName = ThemeFonts["major"];
                            else if (typeface.StartsWith("+mn") && ThemeFonts.ContainsKey("minor"))
                                fontName = ThemeFonts["minor"];
                            else if (!typeface.StartsWith("+"))
                                fontName = typeface;
                        }
                    }

                    Color? tc = ReadSolidFill(rPr, theme);

                    if (tc.HasValue)
                        color = tc.Value;
                }

                sizePt = Math.Max(6f, Math.Min(96f, sizePt));

                StringAlignment alignment = StringAlignment.Near;

                if (pPr != null)
                {
                    string algn = GetAttr(pPr, "algn");

                    if (algn == "ctr")
                        alignment = StringAlignment.Center;
                    else if (algn == "r")
                        alignment = StringAlignment.Far;
                }

                float remainingH = Math.Max(1f, inner.Bottom - y);

                using (Font font = SafeFont(fontName, sizePt, style))
                using (Brush brush = new SolidBrush(color))
                using (StringFormat format = new StringFormat())
                {
                    format.Trimming = StringTrimming.EllipsisCharacter;
                    format.FormatFlags = StringFormatFlags.LineLimit;
                    format.Alignment = alignment;
                    format.LineAlignment = StringAlignment.Near;

                    SizeF measured = g.MeasureString(
                        text.ToString(),
                        font,
                        new SizeF(inner.Width, remainingH),
                        format);

                    float paragraphH = Math.Max(font.Height * 1.15f, measured.Height);

                    RectangleF paragraphRect = new RectangleF(
                        inner.Left,
                        y,
                        inner.Width,
                        Math.Min(remainingH, paragraphH));

                    g.DrawString(
                        text.ToString(),
                        font,
                        brush,
                        paragraphRect,
                        format);

                    y += paragraphH + 2f;
                }

                if (y >= inner.Bottom)
                    break;
            }
        }

        private static bool ContainsCjk(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            for (int i = 0; i < text.Length; i++)
            {
                int c = text[i];
                if ((c >= 0xAC00 && c <= 0xD7AF) ||
                    (c >= 0x3130 && c <= 0x318F) ||
                    (c >= 0x4E00 && c <= 0x9FFF) ||
                    (c >= 0x3040 && c <= 0x30FF))
                    return true;
            }

            return false;
        }

        private static float EmuToRenderPixels(long emu)
        {
            return (float)(emu / EmuPerInch * 144.0);
        }

        private static Font SafeFont(string name, float size)
        {
            return SafeFont(name, size, FontStyle.Regular);
        }

        private static Font SafeFont(string name, float size, FontStyle style)
        {
            try
            {
                return new Font(name, size, style, GraphicsUnit.Point);
            }
            catch
            {
                return new Font("Arial", size, style, GraphicsUnit.Point);
            }
        }

        private static void DrawGraphicFrame(
            ZipArchive zip,
            Graphics g,
            XmlNode frame,
            Dictionary<string, string> rels,
            TransformContext ctx,
            Dictionary<string, Color> theme)
        {
            RectangleF rect;
            if (!TryGetRect(frame, ctx, out rect))
                return;

            XmlNode tbl = FindFirst(frame, "tbl");

            if (tbl != null)
            {
                DrawTable(g, tbl, rect, theme);
                return;
            }

            XmlNode graphicData = FindFirst(frame, "graphicData");
            string uri = graphicData != null ? GetAttr(graphicData, "uri") : "";

            if (uri.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                XmlNode chart = FindFirst(frame, "chart");
                string rid = GetRelationshipId(chart);

                if (!string.IsNullOrEmpty(rid) && rels.ContainsKey(rid))
                {
                    XmlDocument chartDoc = LoadXml(zip, rels[rid]);

                    if (chartDoc != null)
                    {
                        DrawChart(g, chartDoc, rect, theme);
                        return;
                    }
                }

                DrawPlaceholder(g, rect, "Chart");
                return;
            }

            if (uri.IndexOf("diagram", StringComparison.OrdinalIgnoreCase) >= 0 ||
                FindFirst(frame, "relIds") != null)
            {
                XmlNode relIds = FindFirst(frame, "relIds");
                string dm = null;

                if (relIds != null && relIds.Attributes != null)
                {
                    foreach (XmlAttribute a in relIds.Attributes)
                    {
                        if (a.LocalName == "dm")
                        {
                            dm = a.Value;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(dm) && rels.ContainsKey(dm))
                {
                    XmlDocument dataDoc = LoadXml(zip, rels[dm]);
                    if (dataDoc != null)
                    {
                        DrawSmartArt(g, dataDoc, rect, theme);
                        return;
                    }
                }

                DrawPlaceholder(g, rect, "SmartArt / Diagram");
                return;
            }

            XmlNode oleObj = FindFirst(frame, "oleObj");
            if (oleObj != null)
            {
                XmlNode previewBlip = FindFirst(frame, "blip");
                string previewRid = GetRelationshipId(previewBlip);

                if (!string.IsNullOrEmpty(previewRid) && rels.ContainsKey(previewRid))
                {
                    ZipArchiveEntry previewEntry = zip.GetEntry(rels[previewRid]);
                    if (previewEntry != null)
                    {
                        try
                        {
                            using (Stream stream = previewEntry.Open())
                            using (Image image = Image.FromStream(stream))
                            using (Bitmap clone = new Bitmap(image))
                            {
                                g.DrawImage(clone, rect);
                                return;
                            }
                        }
                        catch
                        {
                            if (TryDrawSimpleSvg(previewEntry, g, rect))
                                return;
                        }
                    }
                }

                DrawPlaceholder(g, rect, "Embedded Office object");
                return;
            }

            DrawPlaceholder(g, rect, "Embedded object");
        }

        private sealed class ChartSeriesData
        {
            public string Name;
            public readonly List<string> Categories = new List<string>();
            public readonly List<double> Values = new List<double>();
        }
    }
}
