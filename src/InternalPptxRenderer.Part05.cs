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
            XmlNode resolvedShape = BuildRichInheritedShape(shape);
            DrawRichShapeText(g, resolvedShape, rect, theme, slideNumber);
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
                try
                {
                    return new Font(SystemFonts.MessageBoxFont.FontFamily, size, style, GraphicsUnit.Point);
                }
                catch
                {
                    return new Font("Arial", size, style, GraphicsUnit.Point);
                }
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
                DrawRichTable(g, tbl, rect, theme);
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
            public Color? ExplicitColor;
            public float LineWidth = 2f;
            public DashStyle LineDashStyle = DashStyle.Solid;
            public string MarkerSymbol = "circle";
            public float MarkerSize = 6f;
            public bool MarkerEnabled = true;
            public ChartTrendlineOptions Trendline;
            public ChartErrorBarOptions ErrorBars;
            public string ChartKind;
            public string CategoryAxisId;
            public string ValueAxisId;
            public readonly List<string> Categories = new List<string>();
            public readonly List<double> Values = new List<double>();
            public readonly List<double> BubbleSizes = new List<double>();
            public readonly List<Color?> PointColors = new List<Color?>();
            public readonly Dictionary<int, ChartPointLabelOverride> PointLabelOverrides =
                new Dictionary<int, ChartPointLabelOverride>();
        }
    }
}
