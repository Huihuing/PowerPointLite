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
        private static void DrawPicture(
            ZipArchive zip,
            Graphics g,
            XmlNode pic,
            Dictionary<string, string> rels,
            TransformContext ctx)
        {
            RectangleF rect;
            if (!TryGetRect(pic, ctx, out rect))
                return;

            XmlNode blip = FindFirst(pic, "blip");
            string rid = GetRelationshipId(blip);

            if (string.IsNullOrEmpty(rid) || !rels.ContainsKey(rid))
            {
                DrawPlaceholder(g, rect, "Image");
                return;
            }

            string imagePart = rels[rid];
            ZipArchiveEntry entry = zip.GetEntry(imagePart);

            if (entry == null)
            {
                DrawPlaceholder(g, rect, "Image");
                return;
            }

            try
            {
                using (Stream imageStream = entry.Open())
                using (Image img = Image.FromStream(imageStream))
                using (Bitmap clone = new Bitmap(img))
                {
                    GraphicsState state = g.Save();

                    try
                    {
                        ApplyRotation(g, pic, rect);

                        RectangleF source = new RectangleF(
                            0,
                            0,
                            clone.Width,
                            clone.Height);

                        XmlNode srcRect = FindFirst(pic, "srcRect");

                        if (srcRect != null)
                        {
                            float left = GetLong(srcRect, "l", 0) / 100000f;
                            float top = GetLong(srcRect, "t", 0) / 100000f;
                            float right = GetLong(srcRect, "r", 0) / 100000f;
                            float bottom = GetLong(srcRect, "b", 0) / 100000f;

                            left = Math.Max(0f, Math.Min(0.99f, left));
                            top = Math.Max(0f, Math.Min(0.99f, top));
                            right = Math.Max(0f, Math.Min(0.99f, right));
                            bottom = Math.Max(0f, Math.Min(0.99f, bottom));

                            source = new RectangleF(
                                clone.Width * left,
                                clone.Height * top,
                                Math.Max(1f, clone.Width * (1f - left - right)),
                                Math.Max(1f, clone.Height * (1f - top - bottom)));
                        }

                        XmlNode alphaNode = blip != null ? FindFirst(blip, "alphaModFix") : null;
                        float alpha = 1f;

                        if (alphaNode != null)
                        {
                            long amt = GetLong(alphaNode, "amt", 100000);
                            alpha = Math.Max(0f, Math.Min(1f, amt / 100000f));
                        }

                        if (alpha < 0.999f)
                        {
                            ColorMatrix matrix = new ColorMatrix();
                            matrix.Matrix33 = alpha;

                            using (ImageAttributes attrs = new ImageAttributes())
                            {
                                attrs.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
                                g.DrawImage(
                                    clone,
                                    Rectangle.Round(rect),
                                    source.X, source.Y, source.Width, source.Height,
                                    GraphicsUnit.Pixel,
                                    attrs);
                            }
                        }
                        else
                        {
                            g.DrawImage(
                                clone,
                                Rectangle.Round(rect),
                                source.X,
                                source.Y,
                                source.Width,
                                source.Height,
                                GraphicsUnit.Pixel);
                        }
                    }
                    finally
                    {
                        g.Restore(state);
                    }
                }
            }
            catch
            {
                if (!TryDrawSimpleSvg(entry, g, rect))
                    DrawPlaceholder(g, rect, "Unsupported image");
            }
        }

        private static bool TryDrawSimpleSvg(ZipArchiveEntry entry, Graphics g, RectangleF target)
        {
            if (entry == null ||
                !entry.FullName.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            try
            {
                XmlDocument doc = new XmlDocument();
                using (Stream stream = entry.Open())
                    doc.Load(stream);

                XmlNode svg = doc.DocumentElement;
                if (svg == null || svg.LocalName != "svg")
                    return false;

                float minX = 0f;
                float minY = 0f;
                float vbW = ParseSvgFloat(GetAttr(svg, "width"), 100f);
                float vbH = ParseSvgFloat(GetAttr(svg, "height"), 100f);

                string viewBox = GetAttr(svg, "viewBox");
                if (!string.IsNullOrEmpty(viewBox))
                {
                    string[] pieces = viewBox.Replace(',', ' ').Split(
                        new char[] { ' ', '\t', '\r', '\n' },
                        StringSplitOptions.RemoveEmptyEntries);

                    if (pieces.Length >= 4)
                    {
                        minX = ParseSvgFloat(pieces[0], 0f);
                        minY = ParseSvgFloat(pieces[1], 0f);
                        vbW = Math.Max(0.001f, ParseSvgFloat(pieces[2], vbW));
                        vbH = Math.Max(0.001f, ParseSvgFloat(pieces[3], vbH));
                    }
                }

                float sx = target.Width / Math.Max(0.001f, vbW);
                float sy = target.Height / Math.Max(0.001f, vbH);

                DrawSvgChildren(g, svg, target, minX, minY, sx, sy);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void DrawSvgChildren(
            Graphics g,
            XmlNode parent,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            foreach (XmlNode node in parent.ChildNodes)
            {
                string name = node.LocalName;

                if (name == "g" || name == "svg")
                {
                    DrawSvgChildren(g, node, target, minX, minY, sx, sy);
                    continue;
                }

                Color fill = ReadSvgColor(node, "fill", Color.Black);
                Color stroke = ReadSvgColor(node, "stroke", Color.Transparent);
                float strokeWidth = Math.Max(1f, ParseSvgFloat(GetSvgStyle(node, "stroke-width"), 1f) * Math.Min(sx, sy));

                if (name == "rect")
                {
                    float x = SvgX(target, minX, sx, ParseSvgFloat(GetAttr(node, "x"), 0f));
                    float y = SvgY(target, minY, sy, ParseSvgFloat(GetAttr(node, "y"), 0f));
                    float w = ParseSvgFloat(GetAttr(node, "width"), 0f) * sx;
                    float h = ParseSvgFloat(GetAttr(node, "height"), 0f) * sy;
                    RectangleF r = new RectangleF(x, y, w, h);

                    if (fill.A > 0) using (Brush b = new SolidBrush(fill)) g.FillRectangle(b, r);
                    if (stroke.A > 0) using (Pen p = new Pen(stroke, strokeWidth)) g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height);
                }
                else if (name == "circle")
                {
                    float cx = SvgX(target, minX, sx, ParseSvgFloat(GetAttr(node, "cx"), 0f));
                    float cy = SvgY(target, minY, sy, ParseSvgFloat(GetAttr(node, "cy"), 0f));
                    float rr = ParseSvgFloat(GetAttr(node, "r"), 0f) * Math.Min(sx, sy);
                    RectangleF r = new RectangleF(cx - rr, cy - rr, rr * 2f, rr * 2f);
                    if (fill.A > 0) using (Brush b = new SolidBrush(fill)) g.FillEllipse(b, r);
                    if (stroke.A > 0) using (Pen p = new Pen(stroke, strokeWidth)) g.DrawEllipse(p, r);
                }
                else if (name == "ellipse")
                {
                    float cx = SvgX(target, minX, sx, ParseSvgFloat(GetAttr(node, "cx"), 0f));
                    float cy = SvgY(target, minY, sy, ParseSvgFloat(GetAttr(node, "cy"), 0f));
                    float rx = ParseSvgFloat(GetAttr(node, "rx"), 0f) * sx;
                    float ry = ParseSvgFloat(GetAttr(node, "ry"), 0f) * sy;
                    RectangleF r = new RectangleF(cx - rx, cy - ry, rx * 2f, ry * 2f);
                    if (fill.A > 0) using (Brush b = new SolidBrush(fill)) g.FillEllipse(b, r);
                    if (stroke.A > 0) using (Pen p = new Pen(stroke, strokeWidth)) g.DrawEllipse(p, r);
                }
                else if (name == "line")
                {
                    using (Pen p = new Pen(stroke.A > 0 ? stroke : fill, strokeWidth))
                    {
                        g.DrawLine(p,
                            SvgX(target, minX, sx, ParseSvgFloat(GetAttr(node, "x1"), 0f)),
                            SvgY(target, minY, sy, ParseSvgFloat(GetAttr(node, "y1"), 0f)),
                            SvgX(target, minX, sx, ParseSvgFloat(GetAttr(node, "x2"), 0f)),
                            SvgY(target, minY, sy, ParseSvgFloat(GetAttr(node, "y2"), 0f)));
                    }
                }
                else if (name == "polygon" || name == "polyline")
                {
                    string raw = GetAttr(node, "points");
                    List<PointF> points = ParseSvgPoints(raw, target, minX, minY, sx, sy);

                    if (points.Count >= 2)
                    {
                        if (name == "polygon" && fill.A > 0 && points.Count >= 3)
                            using (Brush b = new SolidBrush(fill)) g.FillPolygon(b, points.ToArray());

                        if (stroke.A > 0 || name == "polyline")
                            using (Pen p = new Pen(stroke.A > 0 ? stroke : fill, strokeWidth))
                            {
                                if (name == "polygon") g.DrawPolygon(p, points.ToArray());
                                else g.DrawLines(p, points.ToArray());
                            }
                    }
                }
                else if (name == "text")
                {
                    float x = SvgX(target, minX, sx, ParseSvgFloat(GetAttr(node, "x"), 0f));
                    float y = SvgY(target, minY, sy, ParseSvgFloat(GetAttr(node, "y"), 0f));
                    float fontSize = Math.Max(6f, ParseSvgFloat(GetSvgStyle(node, "font-size"), 12f) * sy);

                    using (Font font = SafeFont("Arial", fontSize))
                    using (Brush b = new SolidBrush(fill.A > 0 ? fill : Color.Black))
                        g.DrawString(node.InnerText, font, b, x, y - font.Height);
                }
            }
        }

        private static float SvgX(RectangleF target, float minX, float sx, float value)
        {
            return target.Left + (value - minX) * sx;
        }

        private static float SvgY(RectangleF target, float minY, float sy, float value)
        {
            return target.Top + (value - minY) * sy;
        }

        private static string GetSvgStyle(XmlNode node, string key)
        {
            string direct = GetAttr(node, key);
            if (!string.IsNullOrEmpty(direct)) return direct;

            string style = GetAttr(node, "style");
            if (string.IsNullOrEmpty(style)) return null;

            string[] parts = style.Split(';');
            foreach (string part in parts)
            {
                int colon = part.IndexOf(':');
                if (colon <= 0) continue;
                if (part.Substring(0, colon).Trim() == key)
                    return part.Substring(colon + 1).Trim();
            }
            return null;
        }

        private static Color ReadSvgColor(XmlNode node, string key, Color fallback)
        {
            string raw = GetSvgStyle(node, key);
            if (string.IsNullOrEmpty(raw)) return fallback;
            raw = raw.Trim();
            if (raw == "none") return Color.Transparent;

            if (raw.StartsWith("#"))
            {
                Color? c = ParseHexColor(raw.TrimStart('#'));
                if (c.HasValue) return c.Value;
            }

            Color named = Color.FromName(raw);
            return named.A > 0 || string.Equals(raw, "transparent", StringComparison.OrdinalIgnoreCase)
                ? named
                : fallback;
        }

        private static float ParseSvgFloat(string raw, float fallback)
        {
            if (string.IsNullOrEmpty(raw)) return fallback;
            raw = raw.Trim().Replace("px", "").Replace("pt", "");
            float value;
            return float.TryParse(
                raw,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out value) ? value : fallback;
        }

        private static List<PointF> ParseSvgPoints(
            string raw,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            List<PointF> result = new List<PointF>();
            if (string.IsNullOrEmpty(raw)) return result;

            string cleaned = raw.Replace(',', ' ');
            string[] tokens = cleaned.Split(
                new char[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i + 1 < tokens.Length; i += 2)
            {
                float x = ParseSvgFloat(tokens[i], 0f);
                float y = ParseSvgFloat(tokens[i + 1], 0f);
                result.Add(new PointF(
                    target.Left + (x - minX) * sx,
                    target.Top + (y - minY) * sy));
            }

            return result;
        }

        private static void DrawShape(
            ZipArchive zip,
            Graphics g,
            XmlNode shape,
            Dictionary<string, string> rels,
            TransformContext ctx,
            Dictionary<string, Color> theme,
            Dictionary<string, RectangleF> placeholderRects,
            int slideNumber)
        {
            RectangleF rect;

            if (!TryGetRect(shape, ctx, out rect))
            {
                if (!TryGetPlaceholderRect(shape, placeholderRects, out rect))
                    return;
            }

            XmlNode spPr = DirectChild(shape, "spPr");
            string preset = "";

            XmlNode prstGeom = spPr != null ? FindFirst(spPr, "prstGeom") : null;
            if (prstGeom != null)
                preset = GetAttr(prstGeom, "prst");

            bool noFill = spPr != null && FindFirst(spPr, "noFill") != null;
            Brush fillBrush = !noFill && spPr != null
                ? CreateFillBrush(spPr, rect, theme)
                : null;
            Color? line = spPr != null ? ReadLineColor(spPr, theme) : null;

            GraphicsState state = g.Save();

            try
            {
                ApplyRotation(g, shape, rect);

                if (preset == "line")
                {
                    using (Pen pen = CreateLinePen(spPr, theme, line ?? Color.Gray))
                        g.DrawLine(pen, rect.Left, rect.Top, rect.Right, rect.Bottom);
                }
                else
                {
                    using (GraphicsPath path = BuildShapePath(spPr, preset, rect))
                    {
                        DrawShapeVisualEffects(g, spPr, path, rect, theme);

                        bool pictureFill = DrawShapePictureFill(zip, g, spPr, rels, path, rect);

                        if (!pictureFill && fillBrush != null)
                            g.FillPath(fillBrush, path);

                        if (line.HasValue)
                        {
                            using (Pen pen = CreateLinePen(spPr, theme, line.Value))
                                g.DrawPath(pen, path);
                        }
                    }
                }

                DrawShapeText(g, shape, rect, theme, slideNumber);
            }
            finally
            {
                if (fillBrush != null)
                    fillBrush.Dispose();

                g.Restore(state);
            }
        }

        private static bool DrawShapePictureFill(
            ZipArchive zip,
            Graphics g,
            XmlNode spPr,
            Dictionary<string, string> rels,
            GraphicsPath path,
            RectangleF rect)
        {
            if (zip == null || spPr == null || rels == null)
                return false;

            XmlNode blipFill = FindFirst(spPr, "blipFill");
            if (blipFill == null)
                return false;

            XmlNode blip = FindFirst(blipFill, "blip");
            string rid = GetRelationshipId(blip);

            if (string.IsNullOrEmpty(rid) || !rels.ContainsKey(rid))
                return false;

            ZipArchiveEntry entry = zip.GetEntry(rels[rid]);
            if (entry == null)
                return false;

            GraphicsState state = g.Save();

            try
            {
                g.SetClip(path);

                try
                {
                    using (Stream stream = entry.Open())
                    using (Image image = Image.FromStream(stream))
                    using (Bitmap clone = new Bitmap(image))
                    {
                        g.DrawImage(clone, rect);
                        return true;
                    }
                }
                catch
                {
                    return TryDrawSimpleSvg(entry, g, rect);
                }
            }
            finally
            {
                g.Restore(state);
            }
        }

        private static void ApplyRotation(Graphics g, XmlNode node, RectangleF rect)
        {
            XmlNode xfrm = FindFirst(node, "xfrm");
            if (xfrm == null)
                return;

            long raw = GetLong(xfrm, "rot", 0);
            string flipHRaw = GetAttr(xfrm, "flipH");
            string flipVRaw = GetAttr(xfrm, "flipV");
            bool flipH = flipHRaw == "1" || string.Equals(flipHRaw, "true", StringComparison.OrdinalIgnoreCase);
            bool flipV = flipVRaw == "1" || string.Equals(flipVRaw, "true", StringComparison.OrdinalIgnoreCase);

            if (raw == 0 && !flipH && !flipV)
                return;

            float degrees = raw / 60000f;
            float cx = rect.Left + rect.Width / 2f;
            float cy = rect.Top + rect.Height / 2f;

            g.TranslateTransform(cx, cy);

            if (flipH || flipV)
                g.ScaleTransform(flipH ? -1f : 1f, flipV ? -1f : 1f);

            if (raw != 0)
                g.RotateTransform(degrees);

            g.TranslateTransform(-cx, -cy);
        }
    }
}
