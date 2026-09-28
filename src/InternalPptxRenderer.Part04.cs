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
        private static GraphicsPath BuildShapePath(
            XmlNode spPr,
            string preset,
            RectangleF rect)
        {
            XmlNode custom = spPr != null ? FindFirst(spPr, "custGeom") : null;

            if (custom != null)
            {
                GraphicsPath customPath = BuildCustomGeometryPath(custom, rect);
                if (customPath != null && customPath.PointCount > 0)
                    return customPath;

                if (customPath != null)
                    customPath.Dispose();
            }

            return BuildPresetPath(preset, rect);
        }

        private static GraphicsPath BuildCustomGeometryPath(
            XmlNode customGeom,
            RectangleF rect)
        {
            XmlNode pathNode = FindFirst(customGeom, "path");
            if (pathNode == null)
                return null;

            float logicalW = Math.Max(1f, ParseGeometryNumber(GetAttr(pathNode, "w"), 21600f));
            float logicalH = Math.Max(1f, ParseGeometryNumber(GetAttr(pathNode, "h"), 21600f));
            float sx = rect.Width / logicalW;
            float sy = rect.Height / logicalH;

            GraphicsPath path = new GraphicsPath();
            PointF current = new PointF(rect.Left, rect.Top);
            bool hasCurrent = false;

            foreach (XmlNode command in pathNode.ChildNodes)
            {
                if (command.LocalName == "moveTo")
                {
                    XmlNode pt = FindFirst(command, "pt");
                    PointF next;
                    if (TryReadGeometryPoint(pt, rect, sx, sy, out next))
                    {
                        path.StartFigure();
                        current = next;
                        hasCurrent = true;
                    }
                }
                else if (command.LocalName == "lnTo")
                {
                    XmlNode pt = FindFirst(command, "pt");
                    PointF next;
                    if (hasCurrent && TryReadGeometryPoint(pt, rect, sx, sy, out next))
                    {
                        path.AddLine(current, next);
                        current = next;
                    }
                }
                else if (command.LocalName == "cubicBezTo")
                {
                    List<XmlNode> pts = FindAll(command, "pt");
                    if (hasCurrent && pts.Count >= 3)
                    {
                        PointF c1, c2, end;
                        if (TryReadGeometryPoint(pts[0], rect, sx, sy, out c1) &&
                            TryReadGeometryPoint(pts[1], rect, sx, sy, out c2) &&
                            TryReadGeometryPoint(pts[2], rect, sx, sy, out end))
                        {
                            path.AddBezier(current, c1, c2, end);
                            current = end;
                        }
                    }
                }
                else if (command.LocalName == "quadBezTo")
                {
                    List<XmlNode> pts = FindAll(command, "pt");
                    if (hasCurrent && pts.Count >= 2)
                    {
                        PointF control, end;
                        if (TryReadGeometryPoint(pts[0], rect, sx, sy, out control) &&
                            TryReadGeometryPoint(pts[1], rect, sx, sy, out end))
                        {
                            PointF c1 = new PointF(
                                current.X + (control.X - current.X) * 2f / 3f,
                                current.Y + (control.Y - current.Y) * 2f / 3f);
                            PointF c2 = new PointF(
                                end.X + (control.X - end.X) * 2f / 3f,
                                end.Y + (control.Y - end.Y) * 2f / 3f);
                            path.AddBezier(current, c1, c2, end);
                            current = end;
                        }
                    }
                }
                else if (command.LocalName == "close")
                {
                    path.CloseFigure();
                }
            }

            return path;
        }

        private static bool TryReadGeometryPoint(
            XmlNode point,
            RectangleF rect,
            float sx,
            float sy,
            out PointF result)
        {
            result = PointF.Empty;
            if (point == null)
                return false;

            float x;
            float y;

            if (!TryParseGeometryNumber(GetAttr(point, "x"), out x) ||
                !TryParseGeometryNumber(GetAttr(point, "y"), out y))
            {
                return false;
            }

            result = new PointF(rect.Left + x * sx, rect.Top + y * sy);
            return true;
        }

        private static float ParseGeometryNumber(string raw, float fallback)
        {
            float value;
            return TryParseGeometryNumber(raw, out value) ? value : fallback;
        }

        private static bool TryParseGeometryNumber(string raw, out float value)
        {
            value = 0f;
            if (string.IsNullOrEmpty(raw))
                return false;

            return float.TryParse(
                raw,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out value);
        }

        private static GraphicsPath BuildPresetPath(string preset, RectangleF r)
        {
            GraphicsPath path = new GraphicsPath();

            if (string.IsNullOrEmpty(preset) || preset == "rect")
            {
                path.AddRectangle(r);
                return path;
            }

            if (preset == "ellipse")
            {
                path.AddEllipse(r);
                return path;
            }

            if (preset == "roundRect")
            {
                float radius = Math.Max(3f, Math.Min(r.Width, r.Height) * 0.12f);
                float d = radius * 2f;

                path.AddArc(r.Left, r.Top, d, d, 180, 90);
                path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
                path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }

            PointF[] pts = null;

            if (preset == "triangle")
            {
                pts = new PointF[]
                {
                    new PointF(r.Left + r.Width / 2f, r.Top),
                    new PointF(r.Right, r.Bottom),
                    new PointF(r.Left, r.Bottom)
                };
            }
            else if (preset == "rtTriangle")
            {
                pts = new PointF[]
                {
                    new PointF(r.Left, r.Top),
                    new PointF(r.Left, r.Bottom),
                    new PointF(r.Right, r.Bottom)
                };
            }
            else if (preset == "diamond")
            {
                pts = new PointF[]
                {
                    new PointF(r.Left + r.Width / 2f, r.Top),
                    new PointF(r.Right, r.Top + r.Height / 2f),
                    new PointF(r.Left + r.Width / 2f, r.Bottom),
                    new PointF(r.Left, r.Top + r.Height / 2f)
                };
            }
            else if (preset == "parallelogram")
            {
                float dx = r.Width * 0.18f;
                pts = new PointF[]
                {
                    new PointF(r.Left + dx, r.Top),
                    new PointF(r.Right, r.Top),
                    new PointF(r.Right - dx, r.Bottom),
                    new PointF(r.Left, r.Bottom)
                };
            }
            else if (preset == "hexagon")
            {
                float dx = r.Width * 0.24f;
                pts = new PointF[]
                {
                    new PointF(r.Left + dx, r.Top),
                    new PointF(r.Right - dx, r.Top),
                    new PointF(r.Right, r.Top + r.Height / 2f),
                    new PointF(r.Right - dx, r.Bottom),
                    new PointF(r.Left + dx, r.Bottom),
                    new PointF(r.Left, r.Top + r.Height / 2f)
                };
            }
            else if (preset == "pentagon")
            {
                pts = BuildRegularPolygonPoints(r, 5, -90f);
            }
            else if (preset == "octagon")
            {
                pts = BuildRegularPolygonPoints(r, 8, -67.5f);
            }
            else if (preset == "star5")
            {
                pts = BuildStarPoints(r, 5, 0.45f, -90f);
            }
            else if (preset == "plus")
            {
                float x1 = r.Left + r.Width * 0.33f;
                float x2 = r.Left + r.Width * 0.67f;
                float y1 = r.Top + r.Height * 0.33f;
                float y2 = r.Top + r.Height * 0.67f;
                pts = new PointF[]
                {
                    new PointF(x1,r.Top), new PointF(x2,r.Top),
                    new PointF(x2,y1), new PointF(r.Right,y1),
                    new PointF(r.Right,y2), new PointF(x2,y2),
                    new PointF(x2,r.Bottom), new PointF(x1,r.Bottom),
                    new PointF(x1,y2), new PointF(r.Left,y2),
                    new PointF(r.Left,y1), new PointF(x1,y1)
                };
            }
            else if (preset == "trapezoid")
            {
                float dx = r.Width * 0.18f;
                pts = new PointF[]
                {
                    new PointF(r.Left + dx, r.Top),
                    new PointF(r.Right - dx, r.Top),
                    new PointF(r.Right, r.Bottom),
                    new PointF(r.Left, r.Bottom)
                };
            }
            else if (preset == "chevron")
            {
                float dx = r.Width * 0.35f;
                pts = new PointF[]
                {
                    new PointF(r.Left, r.Top),
                    new PointF(r.Right - dx, r.Top),
                    new PointF(r.Right, r.Top + r.Height / 2f),
                    new PointF(r.Right - dx, r.Bottom),
                    new PointF(r.Left, r.Bottom),
                    new PointF(r.Left + dx, r.Top + r.Height / 2f)
                };
            }
            else if (preset == "rightArrow")
            {
                float head = r.Width * 0.35f;
                float bodyTop = r.Top + r.Height * 0.25f;
                float bodyBottom = r.Bottom - r.Height * 0.25f;

                pts = new PointF[]
                {
                    new PointF(r.Left, bodyTop),
                    new PointF(r.Right - head, bodyTop),
                    new PointF(r.Right - head, r.Top),
                    new PointF(r.Right, r.Top + r.Height / 2f),
                    new PointF(r.Right - head, r.Bottom),
                    new PointF(r.Right - head, bodyBottom),
                    new PointF(r.Left, bodyBottom)
                };
            }
            else if (preset == "leftArrow")
            {
                float head = r.Width * 0.35f;
                float bodyTop = r.Top + r.Height * 0.25f;
                float bodyBottom = r.Bottom - r.Height * 0.25f;

                pts = new PointF[]
                {
                    new PointF(r.Right, bodyTop),
                    new PointF(r.Left + head, bodyTop),
                    new PointF(r.Left + head, r.Top),
                    new PointF(r.Left, r.Top + r.Height / 2f),
                    new PointF(r.Left + head, r.Bottom),
                    new PointF(r.Left + head, bodyBottom),
                    new PointF(r.Right, bodyBottom)
                };
            }
            else if (preset == "upArrow")
            {
                float head = r.Height * 0.35f;
                float bodyLeft = r.Left + r.Width * 0.25f;
                float bodyRight = r.Right - r.Width * 0.25f;

                pts = new PointF[]
                {
                    new PointF(r.Left + r.Width / 2f, r.Top),
                    new PointF(r.Right, r.Top + head),
                    new PointF(bodyRight, r.Top + head),
                    new PointF(bodyRight, r.Bottom),
                    new PointF(bodyLeft, r.Bottom),
                    new PointF(bodyLeft, r.Top + head),
                    new PointF(r.Left, r.Top + head)
                };
            }
            else if (preset == "downArrow")
            {
                float head = r.Height * 0.35f;
                float bodyLeft = r.Left + r.Width * 0.25f;
                float bodyRight = r.Right - r.Width * 0.25f;

                pts = new PointF[]
                {
                    new PointF(bodyLeft, r.Top),
                    new PointF(bodyRight, r.Top),
                    new PointF(bodyRight, r.Bottom - head),
                    new PointF(r.Right, r.Bottom - head),
                    new PointF(r.Left + r.Width / 2f, r.Bottom),
                    new PointF(r.Left, r.Bottom - head),
                    new PointF(bodyLeft, r.Bottom - head)
                };
            }

            if (pts != null)
            {
                path.AddPolygon(pts);
                return path;
            }

            path.AddRectangle(r);
            return path;
        }

        private static PointF[] BuildRegularPolygonPoints(
            RectangleF r,
            int sides,
            float startDegrees)
        {
            PointF[] points = new PointF[sides];
            float cx = r.Left + r.Width / 2f;
            float cy = r.Top + r.Height / 2f;
            float rx = r.Width / 2f;
            float ry = r.Height / 2f;

            for (int i = 0; i < sides; i++)
            {
                double angle = (startDegrees + i * 360.0 / sides) * Math.PI / 180.0;
                points[i] = new PointF(
                    cx + (float)Math.Cos(angle) * rx,
                    cy + (float)Math.Sin(angle) * ry);
            }

            return points;
        }

        private static PointF[] BuildStarPoints(
            RectangleF r,
            int points,
            float innerRatio,
            float startDegrees)
        {
            PointF[] result = new PointF[points * 2];
            float cx = r.Left + r.Width / 2f;
            float cy = r.Top + r.Height / 2f;
            float rx = r.Width / 2f;
            float ry = r.Height / 2f;

            for (int i = 0; i < result.Length; i++)
            {
                bool outer = i % 2 == 0;
                float ratio = outer ? 1f : innerRatio;
                double angle = (startDegrees + i * 180.0 / points) * Math.PI / 180.0;
                result[i] = new PointF(
                    cx + (float)Math.Cos(angle) * rx * ratio,
                    cy + (float)Math.Sin(angle) * ry * ratio);
            }

            return result;
        }

        private static void DrawConnector(
            Graphics g,
            XmlNode connector,
            TransformContext ctx,
            Dictionary<string, Color> theme)
        {
            RectangleF rect;
            if (!TryGetRect(connector, ctx, out rect))
                return;

            XmlNode spPr = DirectChild(connector, "spPr");
            Color color = ReadLineColor(connector, theme) ?? Color.Gray;

            using (Pen pen = CreateLinePen(spPr, theme, color))
                g.DrawLine(pen, rect.Left, rect.Top, rect.Right, rect.Bottom);
        }
    }
}
