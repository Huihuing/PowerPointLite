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
            if (customGeom == null)
                return null;

            List<XmlNode> pathNodes =
                FindAll(
                    customGeom,
                    "path");

            if (pathNodes.Count == 0)
                return null;

            GraphicsPath result =
                new GraphicsPath();

            for (int pathIndex = 0;
                 pathIndex < pathNodes.Count;
                 pathIndex++)
            {
                XmlNode pathNode =
                    pathNodes[pathIndex];

                float logicalW =
                    Math.Max(
                        1f,
                        ParseGeometryNumber(
                            GetAttr(
                                pathNode,
                                "w"),
                            21600f));

                float logicalH =
                    Math.Max(
                        1f,
                        ParseGeometryNumber(
                            GetAttr(
                                pathNode,
                                "h"),
                            21600f));

                Dictionary<string, double> guides =
                    BuildGeometryGuideValues(
                        customGeom,
                        logicalW,
                        logicalH);

                float sx =
                    rect.Width /
                    logicalW;
                float sy =
                    rect.Height /
                    logicalH;

                PointF current =
                    new PointF(
                        rect.Left,
                        rect.Top);
                bool hasCurrent =
                    false;

                foreach (XmlNode command in
                    pathNode.ChildNodes)
                {
                    if (command.LocalName == "moveTo")
                    {
                        XmlNode pt =
                            FindFirst(
                                command,
                                "pt");
                        PointF next;

                        if (TryReadGeometryPoint(
                                pt,
                                rect,
                                sx,
                                sy,
                                guides,
                                out next))
                        {
                            result.StartFigure();
                            current = next;
                            hasCurrent = true;
                        }
                    }
                    else if (command.LocalName == "lnTo")
                    {
                        XmlNode pt =
                            FindFirst(
                                command,
                                "pt");
                        PointF next;

                        if (hasCurrent &&
                            TryReadGeometryPoint(
                                pt,
                                rect,
                                sx,
                                sy,
                                guides,
                                out next))
                        {
                            result.AddLine(
                                current,
                                next);
                            current = next;
                        }
                    }
                    else if (command.LocalName == "cubicBezTo")
                    {
                        List<XmlNode> pts =
                            FindAll(
                                command,
                                "pt");

                        if (hasCurrent &&
                            pts.Count >= 3)
                        {
                            PointF c1;
                            PointF c2;
                            PointF finish;

                            if (TryReadGeometryPoint(
                                    pts[0],
                                    rect,
                                    sx,
                                    sy,
                                    guides,
                                    out c1) &&
                                TryReadGeometryPoint(
                                    pts[1],
                                    rect,
                                    sx,
                                    sy,
                                    guides,
                                    out c2) &&
                                TryReadGeometryPoint(
                                    pts[2],
                                    rect,
                                    sx,
                                    sy,
                                    guides,
                                    out finish))
                            {
                                result.AddBezier(
                                    current,
                                    c1,
                                    c2,
                                    finish);
                                current =
                                    finish;
                            }
                        }
                    }
                    else if (command.LocalName == "quadBezTo")
                    {
                        List<XmlNode> pts =
                            FindAll(
                                command,
                                "pt");

                        if (hasCurrent &&
                            pts.Count >= 2)
                        {
                            PointF control;
                            PointF finish;

                            if (TryReadGeometryPoint(
                                    pts[0],
                                    rect,
                                    sx,
                                    sy,
                                    guides,
                                    out control) &&
                                TryReadGeometryPoint(
                                    pts[1],
                                    rect,
                                    sx,
                                    sy,
                                    guides,
                                    out finish))
                            {
                                PointF c1 =
                                    new PointF(
                                        current.X +
                                            (control.X -
                                             current.X) *
                                            2f /
                                            3f,
                                        current.Y +
                                            (control.Y -
                                             current.Y) *
                                            2f /
                                            3f);

                                PointF c2 =
                                    new PointF(
                                        finish.X +
                                            (control.X -
                                             finish.X) *
                                            2f /
                                            3f,
                                        finish.Y +
                                            (control.Y -
                                             finish.Y) *
                                            2f /
                                            3f);

                                result.AddBezier(
                                    current,
                                    c1,
                                    c2,
                                    finish);
                                current =
                                    finish;
                            }
                        }
                    }
                    else if (command.LocalName == "arcTo")
                    {
                        if (!hasCurrent)
                            continue;

                        double wR;
                        double hR;
                        double stAng;
                        double swAng;

                        if (!TryResolveGeometryValue(
                                GetAttr(
                                    command,
                                    "wR"),
                                guides,
                                out wR) ||
                            !TryResolveGeometryValue(
                                GetAttr(
                                    command,
                                    "hR"),
                                guides,
                                out hR) ||
                            !TryResolveGeometryValue(
                                GetAttr(
                                    command,
                                    "stAng"),
                                guides,
                                out stAng) ||
                            !TryResolveGeometryValue(
                                GetAttr(
                                    command,
                                    "swAng"),
                                guides,
                                out swAng))
                        {
                            continue;
                        }

                        float radiusX =
                            Math.Abs(
                                (float)wR *
                                sx);
                        float radiusY =
                            Math.Abs(
                                (float)hR *
                                sy);

                        if (radiusX < 0.01f ||
                            radiusY < 0.01f)
                        {
                            continue;
                        }

                        float startDegrees =
                            (float)(
                                stAng /
                                60000.0);
                        float sweepDegrees =
                            (float)(
                                swAng /
                                60000.0);

                        double startRadians =
                            startDegrees *
                            Math.PI /
                            180.0;

                        float centerX =
                            current.X -
                            radiusX *
                            (float)Math.Cos(
                                startRadians);
                        float centerY =
                            current.Y -
                            radiusY *
                            (float)Math.Sin(
                                startRadians);

                        RectangleF arcRect =
                            new RectangleF(
                                centerX -
                                    radiusX,
                                centerY -
                                    radiusY,
                                radiusX * 2f,
                                radiusY * 2f);

                        result.AddArc(
                            arcRect,
                            startDegrees,
                            sweepDegrees);

                        double endRadians =
                            (startDegrees +
                             sweepDegrees) *
                            Math.PI /
                            180.0;

                        current =
                            new PointF(
                                centerX +
                                    radiusX *
                                    (float)Math.Cos(
                                        endRadians),
                                centerY +
                                    radiusY *
                                    (float)Math.Sin(
                                        endRadians));
                    }
                    else if (command.LocalName == "close")
                    {
                        result.CloseFigure();
                    }
                }
            }

            return result;
        }

        private static Dictionary<string, double>
            BuildGeometryGuideValues(
                XmlNode customGeom,
                float logicalW,
                float logicalH)
        {
            Dictionary<string, double> values =
                new Dictionary<string, double>(
                    StringComparer.Ordinal);

            double w =
                Math.Max(
                    1.0,
                    logicalW);
            double h =
                Math.Max(
                    1.0,
                    logicalH);
            double ss =
                Math.Min(
                    w,
                    h);
            double ls =
                Math.Max(
                    w,
                    h);

            values["l"] = 0.0;
            values["t"] = 0.0;
            values["r"] = w;
            values["b"] = h;
            values["w"] = w;
            values["h"] = h;
            values["hc"] = w / 2.0;
            values["vc"] = h / 2.0;
            values["wd2"] = w / 2.0;
            values["wd3"] = w / 3.0;
            values["wd4"] = w / 4.0;
            values["wd5"] = w / 5.0;
            values["wd6"] = w / 6.0;
            values["wd8"] = w / 8.0;
            values["wd10"] = w / 10.0;
            values["hd2"] = h / 2.0;
            values["hd3"] = h / 3.0;
            values["hd4"] = h / 4.0;
            values["hd5"] = h / 5.0;
            values["hd6"] = h / 6.0;
            values["hd8"] = h / 8.0;
            values["hd10"] = h / 10.0;
            values["ss"] = ss;
            values["ls"] = ls;
            values["ssd2"] = ss / 2.0;
            values["ssd4"] = ss / 4.0;
            values["ssd6"] = ss / 6.0;
            values["ssd8"] = ss / 8.0;
            values["ssd16"] = ss / 16.0;
            values["ssd32"] = ss / 32.0;

            const double fullCircle =
                21600000.0;

            values["cd2"] =
                fullCircle / 2.0;
            values["cd4"] =
                fullCircle / 4.0;
            values["cd8"] =
                fullCircle / 8.0;
            values["3cd4"] =
                fullCircle * 3.0 / 4.0;
            values["3cd8"] =
                fullCircle * 3.0 / 8.0;
            values["5cd8"] =
                fullCircle * 5.0 / 8.0;
            values["7cd8"] =
                fullCircle * 7.0 / 8.0;

            ApplyGeometryGuideList(
                DirectChild(
                    customGeom,
                    "avLst"),
                values);

            ApplyGeometryGuideList(
                DirectChild(
                    customGeom,
                    "gdLst"),
                values);

            return values;
        }

        private static void ApplyGeometryGuideList(
            XmlNode list,
            Dictionary<string, double> values)
        {
            if (list == null ||
                values == null)
            {
                return;
            }

            for (int i = 0;
                 i < list.ChildNodes.Count;
                 i++)
            {
                XmlNode guide =
                    list.ChildNodes[i];

                if (guide.LocalName != "gd")
                    continue;

                string name =
                    GetAttr(
                        guide,
                        "name");
                string formula =
                    GetAttr(
                        guide,
                        "fmla");

                if (string.IsNullOrEmpty(
                        name) ||
                    string.IsNullOrEmpty(
                        formula))
                {
                    continue;
                }

                double value;
                if (TryEvaluateGeometryFormula(
                        formula,
                        values,
                        out value))
                {
                    values[name] =
                        value;
                }
            }
        }

        private static bool TryEvaluateGeometryFormula(
            string formula,
            Dictionary<string, double> values,
            out double result)
        {
            result = 0.0;

            if (string.IsNullOrEmpty(formula))
                return false;

            string[] tokens =
                formula.Split(
                    new char[]
                    {
                        ' ',
                        '\t',
                        '\r',
                        '\n'
                    },
                    StringSplitOptions.RemoveEmptyEntries);

            if (tokens.Length == 0)
                return false;

            string op =
                tokens[0];

            double a;
            double b;
            double c;

            if (op == "val")
            {
                return tokens.Length >= 2 &&
                    TryResolveGeometryValue(
                        tokens[1],
                        values,
                        out result);
            }

            if (op == "abs")
            {
                if (tokens.Length < 2 ||
                    !TryResolveGeometryValue(
                        tokens[1],
                        values,
                        out a))
                {
                    return false;
                }

                result =
                    Math.Abs(a);
                return true;
            }

            if (op == "sqrt")
            {
                if (tokens.Length < 2 ||
                    !TryResolveGeometryValue(
                        tokens[1],
                        values,
                        out a))
                {
                    return false;
                }

                result =
                    Math.Sqrt(
                        Math.Max(
                            0.0,
                            a));
                return true;
            }

            if (tokens.Length < 3 ||
                !TryResolveGeometryValue(
                    tokens[1],
                    values,
                    out a) ||
                !TryResolveGeometryValue(
                    tokens[2],
                    values,
                    out b))
            {
                return false;
            }

            if (op == "max")
            {
                result =
                    Math.Max(
                        a,
                        b);
                return true;
            }

            if (op == "min")
            {
                result =
                    Math.Min(
                        a,
                        b);
                return true;
            }

            if (op == "at2")
            {
                result =
                    Math.Atan2(
                        b,
                        a) *
                    180.0 /
                    Math.PI *
                    60000.0;
                return true;
            }

            if (op == "cos")
            {
                result =
                    a *
                    Math.Cos(
                        GeometryAngleToRadians(
                            b));
                return true;
            }

            if (op == "sin")
            {
                result =
                    a *
                    Math.Sin(
                        GeometryAngleToRadians(
                            b));
                return true;
            }

            if (op == "tan")
            {
                result =
                    a *
                    Math.Tan(
                        GeometryAngleToRadians(
                            b));
                return true;
            }

            if (tokens.Length < 4 ||
                !TryResolveGeometryValue(
                    tokens[3],
                    values,
                    out c))
            {
                return false;
            }

            if (op == "*/")
            {
                result =
                    Math.Abs(c) <
                    0.0000001
                        ? 0.0
                        : a * b / c;
                return true;
            }

            if (op == "+-")
            {
                result =
                    a + b - c;
                return true;
            }

            if (op == "+/")
            {
                result =
                    Math.Abs(c) <
                    0.0000001
                        ? 0.0
                        : (a + b) / c;
                return true;
            }

            if (op == "?:")
            {
                result =
                    a > 0.0
                        ? b
                        : c;
                return true;
            }

            if (op == "mod")
            {
                result =
                    Math.Sqrt(
                        a * a +
                        b * b +
                        c * c);
                return true;
            }

            if (op == "pin")
            {
                double lower =
                    Math.Min(
                        a,
                        c);
                double upper =
                    Math.Max(
                        a,
                        c);

                result =
                    Math.Max(
                        lower,
                        Math.Min(
                            upper,
                            b));
                return true;
            }

            if (op == "cat2")
            {
                result =
                    a *
                    Math.Cos(
                        Math.Atan2(
                            c,
                            b));
                return true;
            }

            if (op == "sat2")
            {
                result =
                    a *
                    Math.Sin(
                        Math.Atan2(
                            c,
                            b));
                return true;
            }

            return false;
        }

        private static double GeometryAngleToRadians(
            double angle)
        {
            return angle /
                60000.0 *
                Math.PI /
                180.0;
        }

        private static bool TryReadGeometryPoint(
            XmlNode point,
            RectangleF rect,
            float sx,
            float sy,
            Dictionary<string, double> guides,
            out PointF result)
        {
            result =
                PointF.Empty;

            if (point == null)
                return false;

            double x;
            double y;

            if (!TryResolveGeometryValue(
                    GetAttr(
                        point,
                        "x"),
                    guides,
                    out x) ||
                !TryResolveGeometryValue(
                    GetAttr(
                        point,
                        "y"),
                    guides,
                    out y))
            {
                return false;
            }

            result =
                new PointF(
                    rect.Left +
                        (float)x *
                        sx,
                    rect.Top +
                        (float)y *
                        sy);
            return true;
        }

        private static bool TryResolveGeometryValue(
            string raw,
            Dictionary<string, double> guides,
            out double value)
        {
            value = 0.0;

            if (string.IsNullOrEmpty(raw))
                return false;

            if (double.TryParse(
                    raw,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out value))
            {
                return true;
            }

            return guides != null &&
                guides.TryGetValue(
                    raw,
                    out value);
        }

        private static float ParseGeometryNumber(
            string raw,
            float fallback)
        {
            float value;

            return TryParseGeometryNumber(
                    raw,
                    out value)
                ? value
                : fallback;
        }

        private static bool TryParseGeometryNumber(
            string raw,
            out float value)
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

            if (string.IsNullOrEmpty(preset) ||
                preset == "rect" ||
                preset == "ellipse" ||
                preset == "roundRect" ||
                preset == "triangle" ||
                preset == "diamond")
            {
                path.Dispose();
                return PresentationRenderPrimitives
                    .CreatePresetShapePath(
                        string.IsNullOrEmpty(preset)
                            ? "rect"
                            : preset,
                        r);
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
