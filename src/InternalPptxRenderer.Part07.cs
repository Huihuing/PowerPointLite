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

        private static Color ApplyColorTransforms(
            Color color,
            XmlNode colorNode)
        {
            if (colorNode == null)
                return color;

            Color current =
                color;

            for (int i = 0;
                 i < colorNode.ChildNodes.Count;
                 i++)
            {
                XmlNode transform =
                    colorNode.ChildNodes[i];

                string name =
                    transform.LocalName;

                if (name == "alpha" ||
                    name == "alphaMod" ||
                    name == "alphaOff")
                {
                    continue;
                }

                if (name == "tint")
                {
                    double factor =
                        Math.Max(
                            0.0,
                            Math.Min(
                                1.0,
                                GetLong(
                                    transform,
                                    "val",
                                    0) /
                                100000.0));

                    current =
                        Color.FromArgb(
                            BlendColorChannel(
                                current.R,
                                255,
                                factor),
                            BlendColorChannel(
                                current.G,
                                255,
                                factor),
                            BlendColorChannel(
                                current.B,
                                255,
                                factor));
                    continue;
                }

                if (name == "shade")
                {
                    double factor =
                        Math.Max(
                            0.0,
                            Math.Min(
                                1.0,
                                GetLong(
                                    transform,
                                    "val",
                                    100000) /
                                100000.0));

                    current =
                        Color.FromArgb(
                            ClampColorByte(
                                current.R *
                                factor),
                            ClampColorByte(
                                current.G *
                                factor),
                            ClampColorByte(
                                current.B *
                                factor));
                    continue;
                }

                if (name == "inv")
                {
                    current =
                        Color.FromArgb(
                            255 -
                                current.R,
                            255 -
                                current.G,
                            255 -
                                current.B);
                    continue;
                }

                if (name == "gray")
                {
                    int gray =
                        ClampColorByte(
                            current.R *
                                0.299 +
                            current.G *
                                0.587 +
                            current.B *
                                0.114);

                    current =
                        Color.FromArgb(
                            gray,
                            gray,
                            gray);
                    continue;
                }

                double hue;
                double saturation;
                double lightness;

                RgbToHsl(
                    current,
                    out hue,
                    out saturation,
                    out lightness);

                if (name == "comp")
                {
                    hue =
                        NormalizeHue(
                            hue +
                            180.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "hue")
                {
                    hue =
                        NormalizeHue(
                            GetLong(
                                transform,
                                "val",
                                0) /
                            60000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "hueOff")
                {
                    hue =
                        NormalizeHue(
                            hue +
                            GetLong(
                                transform,
                                "val",
                                0) /
                            60000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "hueMod")
                {
                    hue =
                        NormalizeHue(
                            hue *
                            GetLong(
                                transform,
                                "val",
                                100000) /
                            100000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "sat")
                {
                    saturation =
                        ClampUnit(
                            GetLong(
                                transform,
                                "val",
                                0) /
                            100000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "satOff")
                {
                    saturation =
                        ClampUnit(
                            saturation +
                            GetLong(
                                transform,
                                "val",
                                0) /
                            100000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "satMod")
                {
                    saturation =
                        ClampUnit(
                            saturation *
                            GetLong(
                                transform,
                                "val",
                                100000) /
                            100000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "lum")
                {
                    lightness =
                        ClampUnit(
                            GetLong(
                                transform,
                                "val",
                                0) /
                            100000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "lumOff")
                {
                    lightness =
                        ClampUnit(
                            lightness +
                            GetLong(
                                transform,
                                "val",
                                0) /
                            100000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "lumMod")
                {
                    lightness =
                        ClampUnit(
                            lightness *
                            GetLong(
                                transform,
                                "val",
                                100000) /
                            100000.0);

                    current =
                        HslToRgb(
                            hue,
                            saturation,
                            lightness);
                    continue;
                }

                if (name == "gamma")
                {
                    current =
                        ApplyGammaColor(
                            current,
                            false);
                    continue;
                }

                if (name == "invGamma")
                {
                    current =
                        ApplyGammaColor(
                            current,
                            true);
                }
            }

            return current;
        }

        private static int BlendColorChannel(
            int from,
            int to,
            double fraction)
        {
            return ClampColorByte(
                from +
                (to -
                 from) *
                fraction);
        }

        private static int ClampColorByte(
            double value)
        {
            return Math.Max(
                0,
                Math.Min(
                    255,
                    (int)Math.Round(
                        value)));
        }

        private static double ClampUnit(
            double value)
        {
            return Math.Max(
                0.0,
                Math.Min(
                    1.0,
                    value));
        }

        private static double NormalizeHue(
            double degrees)
        {
            degrees %=
                360.0;

            if (degrees < 0.0)
                degrees += 360.0;

            return degrees;
        }

        private static void RgbToHsl(
            Color color,
            out double hue,
            out double saturation,
            out double lightness)
        {
            double r =
                color.R /
                255.0;
            double g =
                color.G /
                255.0;
            double b =
                color.B /
                255.0;

            double max =
                Math.Max(
                    r,
                    Math.Max(
                        g,
                        b));
            double min =
                Math.Min(
                    r,
                    Math.Min(
                        g,
                        b));

            lightness =
                (max +
                 min) /
                2.0;

            double delta =
                max -
                min;

            if (delta <
                0.0000001)
            {
                hue = 0.0;
                saturation = 0.0;
                return;
            }

            saturation =
                lightness > 0.5
                    ? delta /
                        (2.0 -
                         max -
                         min)
                    : delta /
                        (max +
                         min);

            double h;

            if (Math.Abs(
                    max -
                    r) <
                0.0000001)
            {
                h =
                    (g -
                     b) /
                    delta +
                    (g < b
                        ? 6.0
                        : 0.0);
            }
            else if (Math.Abs(
                         max -
                         g) <
                     0.0000001)
            {
                h =
                    (b -
                     r) /
                    delta +
                    2.0;
            }
            else
            {
                h =
                    (r -
                     g) /
                    delta +
                    4.0;
            }

            hue =
                NormalizeHue(
                    h *
                    60.0);
        }

        private static Color HslToRgb(
            double hue,
            double saturation,
            double lightness)
        {
            hue =
                NormalizeHue(
                    hue) /
                360.0;
            saturation =
                ClampUnit(
                    saturation);
            lightness =
                ClampUnit(
                    lightness);

            if (saturation <
                0.0000001)
            {
                int gray =
                    ClampColorByte(
                        lightness *
                        255.0);

                return Color.FromArgb(
                    gray,
                    gray,
                    gray);
            }

            double q =
                lightness < 0.5
                    ? lightness *
                        (1.0 +
                         saturation)
                    : lightness +
                        saturation -
                        lightness *
                        saturation;

            double p =
                2.0 *
                lightness -
                q;

            double r =
                HueToRgbChannel(
                    p,
                    q,
                    hue +
                    1.0 /
                    3.0);
            double g =
                HueToRgbChannel(
                    p,
                    q,
                    hue);
            double b =
                HueToRgbChannel(
                    p,
                    q,
                    hue -
                    1.0 /
                    3.0);

            return Color.FromArgb(
                ClampColorByte(
                    r *
                    255.0),
                ClampColorByte(
                    g *
                    255.0),
                ClampColorByte(
                    b *
                    255.0));
        }

        private static double HueToRgbChannel(
            double p,
            double q,
            double t)
        {
            if (t < 0.0)
                t += 1.0;
            if (t > 1.0)
                t -= 1.0;

            if (t <
                1.0 /
                6.0)
            {
                return p +
                    (q -
                     p) *
                    6.0 *
                    t;
            }

            if (t < 0.5)
                return q;

            if (t <
                2.0 /
                3.0)
            {
                return p +
                    (q -
                     p) *
                    (2.0 /
                     3.0 -
                     t) *
                    6.0;
            }

            return p;
        }

        private static Color ApplyGammaColor(
            Color color,
            bool inverse)
        {
            return Color.FromArgb(
                ApplyGammaChannel(
                    color.R,
                    inverse),
                ApplyGammaChannel(
                    color.G,
                    inverse),
                ApplyGammaChannel(
                    color.B,
                    inverse));
        }

        private static int ApplyGammaChannel(
            int channel,
            bool inverse)
        {
            double value =
                channel /
                255.0;

            double result;

            if (inverse)
            {
                result =
                    value <=
                        0.04045
                        ? value /
                            12.92
                        : Math.Pow(
                            (value +
                             0.055) /
                            1.055,
                            2.4);
            }
            else
            {
                result =
                    value <=
                        0.0031308
                        ? value *
                            12.92
                        : 1.055 *
                            Math.Pow(
                                value,
                                1.0 /
                                2.4) -
                            0.055;
            }

            return ClampColorByte(
                ClampUnit(
                    result) *
                255.0);
        }

        private static void DrawShapeVisualEffects(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (g == null ||
                spPr == null ||
                path == null)
            {
                return;
            }

            DrawThreeDimensionalExtrusionApproximation(
                g,
                spPr,
                path,
                rect,
                theme);

            DrawPresetShadowEffect(
                g,
                spPr,
                path,
                rect,
                theme);

            DrawOuterShadowEffect(
                g,
                spPr,
                path,
                rect,
                theme);

            DrawGlowEffect(
                g,
                spPr,
                path,
                theme);

            DrawSoftEdgeApproximation(
                g,
                spPr,
                path,
                theme);

            DrawReflectionEffect(
                g,
                spPr,
                path,
                rect,
                theme);
        }

        private static void DrawThreeDimensionalExtrusionApproximation(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            XmlNode shape3d =
                FindFirst(
                    spPr,
                    "sp3d");

            if (shape3d == null)
                return;

            long extrusionHeight =
                GetLong(
                    shape3d,
                    "extrusionH",
                    0);

            if (extrusionHeight <= 0)
                return;

            float depth =
                Math.Max(
                    1f,
                    Math.Min(
                        28f,
                        EmuEffectToPixels(
                            extrusionHeight)));

            float dx;
            float dy;
            ReadThreeDimensionalProjectionVector(
                spPr,
                out dx,
                out dy);

            float length =
                (float)Math.Sqrt(
                    dx * dx +
                    dy * dy);

            if (length < 0.001f)
            {
                dx = 0.7071f;
                dy = 0.7071f;
            }
            else
            {
                dx /= length;
                dy /= length;
            }

            Color baseColor =
                ReadSolidFill(
                    spPr,
                    theme) ??
                Color.FromArgb(
                    150,
                    150,
                    150);

            XmlNode extrusionColorNode =
                DirectChild(
                    shape3d,
                    "extrusionClr");

            Color extrusionColor =
                extrusionColorNode != null
                    ? ReadColorFromFill(
                        extrusionColorNode,
                        theme) ??
                      DarkenThreeDimensionalColor(
                          baseColor,
                          0.42f)
                    : DarkenThreeDimensionalColor(
                        baseColor,
                        0.42f);

            const int layers = 7;

            for (int i = layers;
                 i >= 1;
                 i--)
            {
                float ratio =
                    i /
                    (float)layers;

                float offsetX =
                    dx *
                    depth *
                    ratio;
                float offsetY =
                    dy *
                    depth *
                    ratio;

                int shade =
                    Math.Max(
                        0,
                        Math.Min(
                            255,
                            (int)Math.Round(
                                36f *
                                (1f -
                                 ratio))));

                Color layerColor =
                    Color.FromArgb(
                        Math.Max(
                            70,
                            Math.Min(
                                235,
                                extrusionColor.A)),
                        Math.Min(
                            255,
                            extrusionColor.R +
                            shade),
                        Math.Min(
                            255,
                            extrusionColor.G +
                            shade),
                        Math.Min(
                            255,
                            extrusionColor.B +
                            shade));

                using (GraphicsPath layer =
                    (GraphicsPath)path.Clone())
                using (Matrix transform =
                    new Matrix())
                using (Brush brush =
                    new SolidBrush(
                        layerColor))
                {
                    transform.Translate(
                        offsetX,
                        offsetY);

                    layer.Transform(
                        transform);

                    g.FillPath(
                        brush,
                        layer);
                }
            }
        }

        private static void ReadThreeDimensionalProjectionVector(
            XmlNode spPr,
            out float dx,
            out float dy)
        {
            dx = 0.7071f;
            dy = 0.7071f;

            XmlNode camera =
                FindFirst(
                    spPr,
                    "camera");

            if (camera == null)
                return;

            XmlNode rotation =
                DirectChild(
                    camera,
                    "rot");

            if (rotation == null)
                return;

            float latitude =
                GetLong(
                    rotation,
                    "lat",
                    0) /
                60000f;

            float longitude =
                GetLong(
                    rotation,
                    "lon",
                    0) /
                60000f;

            float latRadians =
                latitude *
                (float)Math.PI /
                180f;

            float lonRadians =
                longitude *
                (float)Math.PI /
                180f;

            float candidateX =
                -(float)Math.Sin(
                    lonRadians);
            float candidateY =
                (float)Math.Sin(
                    latRadians);

            if (Math.Abs(candidateX) <
                    0.08f &&
                Math.Abs(candidateY) <
                    0.08f)
            {
                return;
            }

            dx = candidateX;
            dy = candidateY;
        }

        private static Color DarkenThreeDimensionalColor(
            Color color,
            float amount)
        {
            amount =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        amount));

            return Color.FromArgb(
                color.A,
                Math.Max(
                    0,
                    (int)Math.Round(
                        color.R *
                        (1f -
                         amount))),
                Math.Max(
                    0,
                    (int)Math.Round(
                        color.G *
                        (1f -
                         amount))),
                Math.Max(
                    0,
                    (int)Math.Round(
                        color.B *
                        (1f -
                         amount))));
        }

        private static void ReadThreeDimensionalLightVector(
            XmlNode spPr,
            out float x,
            out float y)
        {
            x = -1f;
            y = -1f;

            XmlNode lightRig =
                FindFirst(
                    spPr,
                    "lightRig");

            if (lightRig == null)
                return;

            string direction =
                GetAttr(
                    lightRig,
                    "dir");

            if (string.IsNullOrEmpty(direction))
                return;

            direction =
                direction.ToLowerInvariant();

            if (direction == "t")
            {
                x = 0f;
                y = -1f;
            }
            else if (direction == "tr")
            {
                x = 1f;
                y = -1f;
            }
            else if (direction == "r")
            {
                x = 1f;
                y = 0f;
            }
            else if (direction == "br")
            {
                x = 1f;
                y = 1f;
            }
            else if (direction == "b")
            {
                x = 0f;
                y = 1f;
            }
            else if (direction == "bl")
            {
                x = -1f;
                y = 1f;
            }
            else if (direction == "l")
            {
                x = -1f;
                y = 0f;
            }
            else
            {
                x = -1f;
                y = -1f;
            }
        }

        private static void DrawPresetShadowEffect(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (FindFirst(spPr, "outerShdw") != null)
                return;

            XmlNode shadow =
                FindFirst(
                    spPr,
                    "prstShdw");

            if (shadow == null)
                return;

            string preset =
                GetAttr(
                    shadow,
                    "prst");

            long defaultDistance =
                12700L * 4L;

            if (!string.IsNullOrEmpty(preset) &&
                (preset.IndexOf("17", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 preset.IndexOf("18", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 preset.IndexOf("19", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                defaultDistance =
                    12700L * 2L;
            }

            float distance =
                EmuEffectToPixels(
                    GetLong(
                        shadow,
                        "dist",
                        defaultDistance));

            long direction =
                GetLong(
                    shadow,
                    "dir",
                    2700000);

            float radians =
                direction /
                60000f *
                (float)Math.PI /
                180f;

            float dx =
                (float)Math.Cos(
                    radians) *
                distance;
            float dy =
                (float)Math.Sin(
                    radians) *
                distance;

            Color raw =
                ReadColorFromFill(
                    shadow,
                    theme) ??
                Color.Black;

            int alpha =
                raw.A < 255
                    ? raw.A
                    : 68;

            Color color =
                Color.FromArgb(
                    Math.Max(
                        18,
                        Math.Min(
                            160,
                            alpha)),
                    raw.R,
                    raw.G,
                    raw.B);

            using (GraphicsPath shadowPath =
                (GraphicsPath)path.Clone())
            using (Matrix transform =
                new Matrix())
            using (Brush brush =
                new SolidBrush(color))
            {
                transform.Translate(
                    dx,
                    dy);
                shadowPath.Transform(
                    transform);

                g.FillPath(
                    brush,
                    shadowPath);
            }
        }

        private static void DrawShapePostFillEffects(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (g == null ||
                spPr == null ||
                path == null)
            {
                return;
            }

            DrawInnerShadowEffect(
                g,
                spPr,
                path,
                rect,
                theme);

            DrawThreeDimensionalBevelApproximation(
                g,
                spPr,
                path,
                rect,
                theme);
        }

        private static void DrawInnerShadowEffect(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            XmlNode shadow =
                FindFirst(
                    spPr,
                    "innerShdw");

            if (shadow == null)
                return;

            Color raw =
                ReadColorFromFill(
                    shadow,
                    theme) ??
                Color.Black;

            int alpha =
                raw.A < 255
                    ? raw.A
                    : 92;

            float blur =
                Math.Max(
                    1f,
                    Math.Min(
                        32f,
                        EmuEffectToPixels(
                            GetLong(
                                shadow,
                                "blurRad",
                                25400))));

            float distance =
                EmuEffectToPixels(
                    GetLong(
                        shadow,
                        "dist",
                        0));

            long direction =
                GetLong(
                    shadow,
                    "dir",
                    2700000);

            float radians =
                direction /
                60000f *
                (float)Math.PI /
                180f;

            float dx =
                (float)Math.Cos(
                    radians) *
                distance;
            float dy =
                (float)Math.Sin(
                    radians) *
                distance;

            GraphicsState state =
                g.Save();

            try
            {
                g.SetClip(
                    path,
                    CombineMode.Intersect);

                using (GraphicsPath translated =
                    (GraphicsPath)path.Clone())
                using (Matrix transform =
                    new Matrix())
                using (Pen pen =
                    new Pen(
                        Color.FromArgb(
                            Math.Max(
                                24,
                                Math.Min(
                                    180,
                                    alpha)),
                            raw.R,
                            raw.G,
                            raw.B),
                        Math.Max(
                            2f,
                            blur *
                            2.2f)))
                {
                    transform.Translate(
                        dx,
                        dy);
                    translated.Transform(
                        transform);

                    pen.LineJoin =
                        LineJoin.Round;

                    g.DrawPath(
                        pen,
                        translated);
                }
            }
            finally
            {
                g.Restore(state);
            }
        }

        private static void DrawThreeDimensionalBevelApproximation(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            XmlNode shape3d =
                FindFirst(
                    spPr,
                    "sp3d");

            if (shape3d == null)
                return;

            XmlNode topBevel =
                DirectChild(
                    shape3d,
                    "bevelT");
            XmlNode bottomBevel =
                DirectChild(
                    shape3d,
                    "bevelB");

            long topWidthEmu =
                topBevel != null
                    ? GetLong(
                        topBevel,
                        "w",
                        50800)
                    : 0;

            long topHeightEmu =
                topBevel != null
                    ? GetLong(
                        topBevel,
                        "h",
                        topWidthEmu)
                    : 0;

            long bottomWidthEmu =
                bottomBevel != null
                    ? GetLong(
                        bottomBevel,
                        "w",
                        50800)
                    : 0;

            long bottomHeightEmu =
                bottomBevel != null
                    ? GetLong(
                        bottomBevel,
                        "h",
                        bottomWidthEmu)
                    : 0;

            float topSize =
                Math.Max(
                    EmuEffectToPixels(
                        topWidthEmu),
                    EmuEffectToPixels(
                        topHeightEmu));

            float bottomSize =
                Math.Max(
                    EmuEffectToPixels(
                        bottomWidthEmu),
                    EmuEffectToPixels(
                        bottomHeightEmu));

            if (topSize < 0.5f &&
                bottomSize < 0.5f)
            {
                return;
            }

            Color fill =
                ReadSolidFill(
                    spPr,
                    theme) ??
                Color.FromArgb(
                    160,
                    160,
                    160);

            float size =
                Math.Max(
                    1f,
                    Math.Min(
                        18f,
                        Math.Max(
                            topSize,
                            bottomSize)));

            float lightX;
            float lightY;
            ReadThreeDimensionalLightVector(
                spPr,
                out lightX,
                out lightY);

            float lightLength =
                (float)Math.Sqrt(
                    lightX * lightX +
                    lightY * lightY);

            if (lightLength < 0.001f)
            {
                lightX = -0.7071f;
                lightY = -0.7071f;
            }
            else
            {
                lightX /= lightLength;
                lightY /= lightLength;
            }

            GraphicsState state =
                g.Save();

            try
            {
                g.SetClip(
                    path,
                    CombineMode.Intersect);

                if (topSize >= 0.5f)
                {
                    using (GraphicsPath highlightPath =
                        (GraphicsPath)path.Clone())
                    using (Matrix highlightTransform =
                        new Matrix())
                    using (Pen highlight =
                        new Pen(
                            Color.FromArgb(
                                96,
                                255,
                                255,
                                255),
                            Math.Max(
                                1f,
                                size)))
                    {
                        highlightTransform.Translate(
                            lightX *
                            size *
                            0.34f,
                            lightY *
                            size *
                            0.34f);
                        highlightPath.Transform(
                            highlightTransform);
                        highlight.LineJoin =
                            LineJoin.Round;
                        g.DrawPath(
                            highlight,
                            highlightPath);
                    }
                }

                if (bottomSize >= 0.5f ||
                    topSize >= 0.5f)
                {
                    using (GraphicsPath shadePath =
                        (GraphicsPath)path.Clone())
                    using (Matrix shadeTransform =
                        new Matrix())
                    using (Pen shade =
                        new Pen(
                            Color.FromArgb(
                                92,
                                Math.Max(
                                    0,
                                    fill.R -
                                    70),
                                Math.Max(
                                    0,
                                    fill.G -
                                    70),
                                Math.Max(
                                    0,
                                    fill.B -
                                    70)),
                            Math.Max(
                                1f,
                                size)))
                    {
                        shadeTransform.Translate(
                            -lightX *
                            size *
                            0.34f,
                            -lightY *
                            size *
                            0.34f);
                        shadePath.Transform(
                            shadeTransform);
                        shade.LineJoin =
                            LineJoin.Round;
                        g.DrawPath(
                            shade,
                            shadePath);
                    }
                }
            }
            finally
            {
                g.Restore(state);
            }
        }

        private static void DrawOuterShadowEffect(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            XmlNode shadow =
                FindFirst(
                    spPr,
                    "outerShdw");

            if (shadow == null)
                return;

            Color shadowColor =
                Color.FromArgb(
                    85,
                    0,
                    0,
                    0);

            Color? rawColor =
                ReadColorFromFill(
                    shadow,
                    theme);

            if (rawColor.HasValue)
            {
                int alpha =
                    rawColor.Value.A < 255
                        ? rawColor.Value.A
                        : 80;

                shadowColor =
                    Color.FromArgb(
                        alpha,
                        rawColor.Value.R,
                        rawColor.Value.G,
                        rawColor.Value.B);
            }

            long distEmu =
                GetLong(
                    shadow,
                    "dist",
                    0);
            long dirRaw =
                GetLong(
                    shadow,
                    "dir",
                    2700000);

            float distPx =
                EmuEffectToPixels(
                    distEmu);

            float angle =
                dirRaw /
                60000f *
                (float)Math.PI /
                180f;

            float dx =
                (float)Math.Cos(
                    angle) *
                distPx;
            float dy =
                (float)Math.Sin(
                    angle) *
                distPx;

            using (GraphicsPath shadowPath =
                (GraphicsPath)path.Clone())
            using (Matrix matrix =
                new Matrix())
            using (Brush brush =
                new SolidBrush(
                    shadowColor))
            {
                matrix.Translate(
                    dx,
                    dy);
                shadowPath.Transform(
                    matrix);
                g.FillPath(
                    brush,
                    shadowPath);
            }
        }

        private static void DrawGlowEffect(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            Dictionary<string, Color> theme)
        {
            XmlNode glow =
                FindFirst(
                    spPr,
                    "glow");

            if (glow == null)
                return;

            float radius =
                Math.Max(
                    1f,
                    Math.Min(
                        80f,
                        EmuEffectToPixels(
                            GetLong(
                                glow,
                                "rad",
                                0))));

            if (radius <= 1f)
                return;

            Color baseColor =
                ReadColorFromFill(
                    glow,
                    theme) ??
                Color.FromArgb(
                    100,
                    90,
                    150,
                    255);

            int baseAlpha =
                baseColor.A < 255
                    ? baseColor.A
                    : 110;

            const int layers = 8;

            for (int i = layers;
                 i >= 1;
                 i--)
            {
                float ratio =
                    i /
                    (float)layers;

                float width =
                    Math.Max(
                        1f,
                        radius *
                        2f *
                        ratio);

                int alpha =
                    Math.Max(
                        3,
                        Math.Min(
                            180,
                            (int)Math.Round(
                                baseAlpha *
                                (1f -
                                 ratio *
                                 0.72f) /
                                layers *
                                2.6f)));

                using (Pen pen =
                    new Pen(
                        Color.FromArgb(
                            alpha,
                            baseColor.R,
                            baseColor.G,
                            baseColor.B),
                        width))
                {
                    pen.LineJoin =
                        LineJoin.Round;

                    g.DrawPath(
                        pen,
                        path);
                }
            }
        }

        private static void DrawSoftEdgeApproximation(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            Dictionary<string, Color> theme)
        {
            XmlNode softEdge =
                FindFirst(
                    spPr,
                    "softEdge");

            if (softEdge == null)
                return;

            float radius =
                Math.Max(
                    0f,
                    Math.Min(
                        48f,
                        EmuEffectToPixels(
                            GetLong(
                                softEdge,
                                "rad",
                                0))));

            if (radius < 1f)
                return;

            Color baseColor =
                ReadSolidFill(
                    spPr,
                    theme) ??
                Color.Gray;

            const int layers = 6;

            for (int i = layers;
                 i >= 1;
                 i--)
            {
                float ratio =
                    i /
                    (float)layers;

                int alpha =
                    Math.Max(
                        2,
                        (int)Math.Round(
                            34f *
                            (1f -
                             ratio *
                             0.70f)));

                using (Pen pen =
                    new Pen(
                        Color.FromArgb(
                            alpha,
                            baseColor.R,
                            baseColor.G,
                            baseColor.B),
                        Math.Max(
                            1f,
                            radius *
                            2f *
                            ratio)))
                {
                    pen.LineJoin =
                        LineJoin.Round;

                    g.DrawPath(
                        pen,
                        path);
                }
            }
        }

        private static void DrawReflectionEffect(
            Graphics g,
            XmlNode spPr,
            GraphicsPath path,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            XmlNode reflection =
                FindFirst(
                    spPr,
                    "reflection");

            if (reflection == null)
                return;

            Color sourceColor =
                ReadSolidFill(
                    spPr,
                    theme) ??
                Color.FromArgb(
                    120,
                    120,
                    120);

            float scaleY =
                GetLong(
                    reflection,
                    "sy",
                    100000) /
                100000f;

            scaleY =
                Math.Max(
                    0.05f,
                    Math.Min(
                        2f,
                        Math.Abs(
                            scaleY)));

            float distance =
                EmuEffectToPixels(
                    GetLong(
                        reflection,
                        "dist",
                        0));

            int startAlpha =
                (int)Math.Round(
                    255.0 *
                    Math.Max(
                        0.0,
                        Math.Min(
                            1.0,
                            GetLong(
                                reflection,
                                "stA",
                                52000) /
                            100000.0)));

            int endAlpha =
                (int)Math.Round(
                    255.0 *
                    Math.Max(
                        0.0,
                        Math.Min(
                            1.0,
                            GetLong(
                                reflection,
                                "endA",
                                0) /
                            100000.0)));

            using (GraphicsPath reflected =
                (GraphicsPath)path.Clone())
            using (Matrix mirror =
                new Matrix(
                    1f,
                    0f,
                    0f,
                    -scaleY,
                    0f,
                    rect.Bottom *
                        (1f +
                         scaleY) +
                    distance))
            {
                reflected.Transform(
                    mirror);

                RectangleF bounds =
                    reflected.GetBounds();

                if (bounds.Width <= 0f ||
                    bounds.Height <= 0f)
                {
                    return;
                }

                using (LinearGradientBrush brush =
                    new LinearGradientBrush(
                        new PointF(
                            bounds.Left,
                            bounds.Top),
                        new PointF(
                            bounds.Left,
                            bounds.Bottom),
                        Color.FromArgb(
                            startAlpha,
                            sourceColor.R,
                            sourceColor.G,
                            sourceColor.B),
                        Color.FromArgb(
                            endAlpha,
                            sourceColor.R,
                            sourceColor.G,
                            sourceColor.B)))
                {
                    g.FillPath(
                        brush,
                        reflected);
                }
            }
        }

        private static float EmuEffectToPixels(
            long emu)
        {
            return (float)(
                emu /
                EmuPerInch *
                144.0);
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
