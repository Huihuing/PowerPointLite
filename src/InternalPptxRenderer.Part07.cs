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
        private static void DrawTable(
            Graphics g,
            XmlNode tbl,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            List<XmlNode> rows = FindAll(tbl, "tr");
            if (rows.Count == 0)
                return;

            int cols = 0;
            foreach (XmlNode row in rows)
                cols = Math.Max(cols, FindAll(row, "tc").Count);

            if (cols == 0)
                return;

            float rowH = rect.Height / rows.Count;
            float colW = rect.Width / cols;

            using (Pen pen = new Pen(Color.Gray, 1f))
            {
                for (int r = 0; r < rows.Count; r++)
                {
                    List<XmlNode> cells = FindAll(rows[r], "tc");

                    for (int c = 0; c < cells.Count; c++)
                    {
                        RectangleF cell = new RectangleF(
                            rect.X + c * colW,
                            rect.Y + r * rowH,
                            colW,
                            rowH);

                        Color? fill = ReadSolidFill(cells[c], theme);

                        if (fill.HasValue)
                        {
                            using (Brush b = new SolidBrush(fill.Value))
                                g.FillRectangle(b, cell);
                        }

                        g.DrawRectangle(pen, cell.X, cell.Y, cell.Width, cell.Height);

                        StringBuilder sb = new StringBuilder();
                        foreach (XmlNode t in FindAll(cells[c], "t"))
                            sb.Append(t.InnerText);

                        if (sb.Length > 0)
                        {
                            using (Font font = SafeFont("Arial", 11f))
                            using (Brush brush = new SolidBrush(Color.Black))
                            {
                                RectangleF textRect = new RectangleF(
                                    cell.X + 3, cell.Y + 2,
                                    Math.Max(1, cell.Width - 6),
                                    Math.Max(1, cell.Height - 4));

                                g.DrawString(sb.ToString(), font, brush, textRect);
                            }
                        }
                    }
                }
            }
        }

        private static void DrawPlaceholder(Graphics g, RectangleF rect, string label)
        {
            using (Brush b = new SolidBrush(Color.FromArgb(235, 235, 235)))
            using (Pen p = new Pen(Color.FromArgb(150, 150, 150), 1f))
            using (Font f = SafeFont("Arial", 12f))
            using (Brush tb = new SolidBrush(Color.DimGray))
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;

                g.FillRectangle(b, rect);
                g.DrawRectangle(p, rect.X, rect.Y, rect.Width, rect.Height);
                g.DrawString(label, f, tb, rect, sf);
            }
        }

        private static bool TryGetRect(
            XmlNode node,
            TransformContext ctx,
            out RectangleF rect)
        {
            rect = RectangleF.Empty;

            XmlNode xfrm = null;

            XmlNode spPr = DirectChild(node, "spPr");
            if (spPr != null)
                xfrm = DirectChild(spPr, "xfrm");

            if (xfrm == null)
                xfrm = DirectChild(node, "xfrm");

            if (xfrm == null)
            {
                XmlNode grpPr = DirectChild(node, "grpSpPr");
                if (grpPr != null)
                    xfrm = DirectChild(grpPr, "xfrm");
            }

            if (xfrm == null)
                return false;

            XmlNode off = DirectChild(xfrm, "off");
            XmlNode ext = DirectChild(xfrm, "ext");

            if (off == null || ext == null)
                return false;

            long x = GetLong(off, "x", 0);
            long y = GetLong(off, "y", 0);
            long cx = GetLong(ext, "cx", 0);
            long cy = GetLong(ext, "cy", 0);

            rect = new RectangleF(
                ctx.X(x),
                ctx.Y(y),
                Math.Max(1f, ctx.W(cx)),
                Math.Max(1f, ctx.H(cy)));

            return true;
        }

        private static void DrawBackgroundLayer(
            ZipArchive zip,
            Graphics g,
            XmlDocument doc,
            Dictionary<string, string> rels,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (doc == null)
                return;

            XmlNode bg = FindFirst(doc, "bg");
            if (bg == null)
                return;

            XmlNode blip = FindFirst(bg, "blip");
            string rid = GetRelationshipId(blip);

            if (!string.IsNullOrEmpty(rid) && rels != null && rels.ContainsKey(rid))
            {
                ZipArchiveEntry entry = zip.GetEntry(rels[rid]);

                if (entry != null)
                {
                    try
                    {
                        using (Stream stream = entry.Open())
                        using (Image img = Image.FromStream(stream))
                        using (Bitmap clone = new Bitmap(img))
                        {
                            g.DrawImage(clone, rect);
                            return;
                        }
                    }
                    catch
                    {
                        if (TryDrawSimpleSvg(entry, g, rect))
                            return;
                    }
                }
            }

            Brush fill = CreateFillBrush(bg, rect, theme);
            if (fill != null)
            {
                using (fill)
                    g.FillRectangle(fill, rect);
                return;
            }

            XmlNode bgRef = FindFirst(bg, "bgRef");
            if (bgRef != null)
            {
                Color? color = ReadColorFromFill(bgRef, theme);
                if (color.HasValue)
                {
                    using (Brush brush = new SolidBrush(color.Value))
                        g.FillRectangle(brush, rect);
                }
            }
        }

        private static Color? ReadBackground(
            XmlDocument doc,
            Dictionary<string, Color> theme)
        {
            if (doc == null)
                return null;

            XmlNode bg = FindFirst(doc, "bg");
            if (bg == null)
                return null;

            return ReadSolidFill(bg, theme);
        }

        private static Pen CreateLinePen(
            XmlNode spPr,
            Dictionary<string, Color> theme,
            Color fallback)
        {
            XmlNode ln = spPr != null ? FindFirst(spPr, "ln") : null;
            Color color = ln != null ? (ReadSolidFill(ln, theme) ?? fallback) : fallback;
            float width = 1.5f;

            if (ln != null)
            {
                long rawWidth = GetLong(ln, "w", 0);
                if (rawWidth > 0)
                    width = Math.Max(1f, (float)(rawWidth / 12700.0 * 2.0));
            }

            Pen pen = new Pen(color, width);
            pen.LineJoin = LineJoin.Round;

            if (ln != null)
            {
                XmlNode dash = FindFirst(ln, "prstDash");
                string dashVal = dash != null ? GetAttr(dash, "val") : null;

                if (dashVal == "dash" || dashVal == "lgDash")
                    pen.DashStyle = DashStyle.Dash;
                else if (dashVal == "dot" || dashVal == "sysDot")
                    pen.DashStyle = DashStyle.Dot;
                else if (dashVal == "dashDot" || dashVal == "sysDashDot")
                    pen.DashStyle = DashStyle.DashDot;

                XmlNode head = FindFirst(ln, "headEnd");
                XmlNode tail = FindFirst(ln, "tailEnd");

                if (head != null && GetAttr(head, "type") != "none")
                    pen.StartCap = LineCap.ArrowAnchor;

                if (tail != null && GetAttr(tail, "type") != "none")
                    pen.EndCap = LineCap.ArrowAnchor;
            }

            return pen;
        }

        private static Color? ReadLineColor(
            XmlNode node,
            Dictionary<string, Color> theme)
        {
            XmlNode ln = FindFirst(node, "ln");
            if (ln == null)
                return null;

            return ReadSolidFill(ln, theme);
        }

        private static Brush CreateFillBrush(
            XmlNode node,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (node == null)
                return null;

            XmlNode solid = FindFirst(node, "solidFill");
            if (solid != null)
            {
                Color? color = ReadColorFromFill(solid, theme);
                if (color.HasValue)
                    return new SolidBrush(color.Value);
            }

            XmlNode pattern = FindFirst(node, "pattFill");
            if (pattern != null)
            {
                XmlNode fg = DirectChild(pattern, "fgClr");
                XmlNode bg = DirectChild(pattern, "bgClr");
                Color foreground = ReadColorFromFill(fg, theme) ?? Color.Gray;
                Color background = ReadColorFromFill(bg, theme) ?? Color.White;
                string prst = GetAttr(pattern, "prst") ?? "pct20";
                HatchStyle hatch = HatchStyle.Percent20;

                if (prst.IndexOf("horz", StringComparison.OrdinalIgnoreCase) >= 0) hatch = HatchStyle.Horizontal;
                else if (prst.IndexOf("vert", StringComparison.OrdinalIgnoreCase) >= 0) hatch = HatchStyle.Vertical;
                else if (prst.IndexOf("diagCross", StringComparison.OrdinalIgnoreCase) >= 0) hatch = HatchStyle.DiagonalCross;
                else if (prst.IndexOf("cross", StringComparison.OrdinalIgnoreCase) >= 0) hatch = HatchStyle.Cross;
                else if (prst.IndexOf("dnDiag", StringComparison.OrdinalIgnoreCase) >= 0) hatch = HatchStyle.BackwardDiagonal;
                else if (prst.IndexOf("upDiag", StringComparison.OrdinalIgnoreCase) >= 0) hatch = HatchStyle.ForwardDiagonal;

                return new HatchBrush(hatch, foreground, background);
            }

            XmlNode grad = FindFirst(node, "gradFill");
            if (grad != null)
            {
                List<XmlNode> stops = FindAll(grad, "gs");
                Color first = Color.White;
                Color last = Color.LightGray;

                if (stops.Count > 0)
                {
                    Color? c = ReadColorFromFill(stops[0], theme);
                    if (c.HasValue) first = c.Value;

                    c = ReadColorFromFill(stops[stops.Count - 1], theme);
                    if (c.HasValue) last = c.Value;
                }

                float angle = 0f;
                XmlNode lin = FindFirst(grad, "lin");

                if (lin != null)
                {
                    long raw = GetLong(lin, "ang", 0);
                    angle = raw / 60000f;
                }

                if (rect.Width > 0 && rect.Height > 0)
                    return new LinearGradientBrush(rect, first, last, angle);
            }

            return null;
        }

        private static Color? ReadColorFromFill(
            XmlNode node,
            Dictionary<string, Color> theme)
        {
            if (node == null)
                return null;

            XmlNode colorNode = DirectChild(node, "srgbClr");
            Color? color = null;

            if (colorNode != null)
                color = ParseHexColor(GetAttr(colorNode, "val"));

            if (!color.HasValue)
            {
                colorNode = DirectChild(node, "sysClr");
                if (colorNode != null)
                    color = ParseHexColor(GetAttr(colorNode, "lastClr"));
            }

            if (!color.HasValue)
            {
                colorNode = DirectChild(node, "schemeClr");
                if (colorNode != null)
                {
                    string key = GetAttr(colorNode, "val");

                    if (key == "tx1") key = "dk1";
                    else if (key == "bg1") key = "lt1";
                    else if (key == "tx2") key = "dk2";
                    else if (key == "bg2") key = "lt2";

                    if (!string.IsNullOrEmpty(key) && theme != null && theme.ContainsKey(key))
                        color = theme[key];
                }
            }

            if (!color.HasValue)
            {
                colorNode = DirectChild(node, "prstClr");
                if (colorNode != null)
                {
                    string preset = GetAttr(colorNode, "val");
                    Color named = Color.FromName(preset ?? "");
                    if (named.A > 0) color = named;
                }
            }

            if (!color.HasValue)
                return null;

            int alpha = color.Value.A;

            if (colorNode != null)
            {
                XmlNode alphaNode = FindFirst(colorNode, "alpha");
                if (alphaNode != null)
                {
                    long raw = GetLong(alphaNode, "val", 100000);
                    alpha = Math.Max(0, Math.Min(255, (int)Math.Round(255 * raw / 100000.0)));
                }

                XmlNode alphaMod = FindFirst(colorNode, "alphaMod");
                if (alphaMod != null)
                {
                    long raw = GetLong(alphaMod, "val", 100000);
                    alpha = Math.Max(0, Math.Min(255, (int)Math.Round(alpha * raw / 100000.0)));
                }
            }

            Color transformed = ApplyColorTransforms(color.Value, colorNode);
            return Color.FromArgb(alpha, transformed.R, transformed.G, transformed.B);
        }

        private static Color ApplyColorTransforms(Color color, XmlNode colorNode)
        {
            if (colorNode == null) return color;

            double r = color.R;
            double g = color.G;
            double b = color.B;

            XmlNode tint = DirectChild(colorNode, "tint");
            if (tint != null)
            {
                double f = Math.Max(0, Math.Min(1, GetLong(tint, "val", 0) / 100000.0));
                r = r + (255 - r) * f;
                g = g + (255 - g) * f;
                b = b + (255 - b) * f;
            }

            XmlNode shade = DirectChild(colorNode, "shade");
            if (shade != null)
            {
                double f = Math.Max(0, Math.Min(1, GetLong(shade, "val", 100000) / 100000.0));
                r *= f; g *= f; b *= f;
            }

            XmlNode lumMod = DirectChild(colorNode, "lumMod");
            if (lumMod != null)
            {
                double f = GetLong(lumMod, "val", 100000) / 100000.0;
                r *= f; g *= f; b *= f;
            }

            XmlNode lumOff = DirectChild(colorNode, "lumOff");
            if (lumOff != null)
            {
                double f = GetLong(lumOff, "val", 0) / 100000.0 * 255.0;
                r += f; g += f; b += f;
            }

            return Color.FromArgb(
                Math.Max(0, Math.Min(255, (int)Math.Round(r))),
                Math.Max(0, Math.Min(255, (int)Math.Round(g))),
                Math.Max(0, Math.Min(255, (int)Math.Round(b))));
        }

        private static void DrawOuterShadow(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (spPr == null || path == null)
                return;

            XmlNode shadow = FindFirst(spPr, "outerShdw");
            if (shadow == null)
                return;

            Color shadowColor = Color.FromArgb(85, 0, 0, 0);
            Color? rawColor = ReadColorFromFill(shadow, theme);

            if (rawColor.HasValue)
            {
                int alpha = rawColor.Value.A < 255 ? rawColor.Value.A : 80;
                shadowColor = Color.FromArgb(alpha, rawColor.Value.R, rawColor.Value.G, rawColor.Value.B);
            }

            long distEmu = GetLong(shadow, "dist", 0);
            long dirRaw = GetLong(shadow, "dir", 2700000);
            float distPx = (float)(distEmu / EmuPerInch * 144.0);
            float angle = dirRaw / 60000f * (float)Math.PI / 180f;
            float dx = (float)Math.Cos(angle) * distPx;
            float dy = (float)Math.Sin(angle) * distPx;

            using (GraphicsPath shadowPath = (GraphicsPath)path.Clone())
            using (Matrix m = new Matrix())
            using (Brush brush = new SolidBrush(shadowColor))
            {
                m.Translate(dx, dy);
                shadowPath.Transform(m);
                g.FillPath(brush, shadowPath);
            }
        }

        private static Color? ReadSolidFill(
            XmlNode node,
            Dictionary<string, Color> theme)
        {
            if (node == null)
                return null;

            XmlNode solid = FindFirst(node, "solidFill");
            return solid != null ? ReadColorFromFill(solid, theme) : (Color?)null;
        }

        private static Color? ParseHexColor(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;

            value = value.Trim().TrimStart('#');

            if (value.Length != 6)
                return null;

            int r, g, b;

            if (int.TryParse(value.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out r) &&
                int.TryParse(value.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out g) &&
                int.TryParse(value.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out b))
            {
                return Color.FromArgb(r, g, b);
            }

            return null;
        }

        private static void LoadThemeFonts(ZipArchive zip)
        {
            ThemeFonts.Clear();
            ThemeFonts["major"] = "Arial";
            ThemeFonts["minor"] = "Arial";
            ThemeFonts["majorEA"] = "Arial";
            ThemeFonts["minorEA"] = "Arial";

            XmlDocument theme = LoadXml(zip, "ppt/theme/theme1.xml");
            if (theme == null)
                return;

            LoadThemeFontGroup(FindFirst(theme, "majorFont"), "major", "majorEA");
            LoadThemeFontGroup(FindFirst(theme, "minorFont"), "minor", "minorEA");
        }

        private static void LoadThemeFontGroup(
            XmlNode group,
            string latinKey,
            string eastAsiaKey)
        {
            if (group == null)
                return;

            XmlNode latin = DirectChild(group, "latin");
            XmlNode eastAsia = DirectChild(group, "ea");

            string latinFace = latin != null ? GetAttr(latin, "typeface") : null;
            string eastFace = eastAsia != null ? GetAttr(eastAsia, "typeface") : null;

            foreach (XmlNode font in FindAll(group, "font"))
            {
                string script = GetAttr(font, "script");
                string face = GetAttr(font, "typeface");

                if (string.IsNullOrEmpty(eastFace) &&
                    (script == "Hang" || script == "Hans" || script == "Hant" || script == "Jpan"))
                {
                    eastFace = face;
                }
            }

            if (!string.IsNullOrEmpty(latinFace)) ThemeFonts[latinKey] = latinFace;
            if (!string.IsNullOrEmpty(eastFace)) ThemeFonts[eastAsiaKey] = eastFace;
            else ThemeFonts[eastAsiaKey] = ThemeFonts[latinKey];
        }
    }
}
