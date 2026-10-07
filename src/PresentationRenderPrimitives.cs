using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text;

namespace PptxViewer
{
    internal static class PresentationRenderPrimitives
    {
        public static Font SafeFont(
            string family,
            float size,
            FontStyle style)
        {
            float safeSize = size;

            if (float.IsNaN(safeSize) ||
                float.IsInfinity(safeSize) ||
                safeSize < 1f)
            {
                safeSize = 1f;
            }

            string name =
                string.IsNullOrEmpty(family)
                    ? "Arial"
                    : family;

            try
            {
                return new Font(
                    name,
                    safeSize,
                    style,
                    GraphicsUnit.Point);
            }
            catch
            {
                try
                {
                    return new Font(
                        SystemFonts.MessageBoxFont.FontFamily,
                        safeSize,
                        style,
                        GraphicsUnit.Point);
                }
                catch
                {
                    return new Font(
                        "Arial",
                        safeSize,
                        style,
                        GraphicsUnit.Point);
                }
            }
        }

        public static Color ParseHexColor(
            string value,
            Color fallback)
        {
            string candidate =
                (value ?? string.Empty)
                .Trim()
                .TrimStart('#');
            int parsed;

            if (candidate.Length == 6 &&
                int.TryParse(
                    candidate,
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return Color.FromArgb(
                    (parsed >> 16) & 0xFF,
                    (parsed >> 8) & 0xFF,
                    parsed & 0xFF);
            }

            return fallback;
        }

        public static int ClampCropValue(
            int value)
        {
            return Math.Max(
                0,
                Math.Min(
                    99999,
                    value));
        }

        public static void NormalizeCrop(
            int left,
            int top,
            int right,
            int bottom,
            out int normalizedLeft,
            out int normalizedTop,
            out int normalizedRight,
            out int normalizedBottom)
        {
            normalizedLeft = ClampCropValue(left);
            normalizedTop = ClampCropValue(top);
            normalizedRight = ClampCropValue(right);
            normalizedBottom = ClampCropValue(bottom);

            NormalizeCropPair(
                ref normalizedLeft,
                ref normalizedRight);
            NormalizeCropPair(
                ref normalizedTop,
                ref normalizedBottom);
        }

        public static RectangleF CalculateImageSourceRectangle(
            int pixelWidth,
            int pixelHeight,
            int cropLeft,
            int cropTop,
            int cropRight,
            int cropBottom)
        {
            int safeWidth = Math.Max(1, pixelWidth);
            int safeHeight = Math.Max(1, pixelHeight);
            int left;
            int top;
            int right;
            int bottom;

            NormalizeCrop(
                cropLeft,
                cropTop,
                cropRight,
                cropBottom,
                out left,
                out top,
                out right,
                out bottom);

            float leftRatio = left / 100000f;
            float topRatio = top / 100000f;
            float rightRatio = right / 100000f;
            float bottomRatio = bottom / 100000f;

            return new RectangleF(
                safeWidth * leftRatio,
                safeHeight * topRatio,
                Math.Max(
                    1f,
                    safeWidth *
                    (1f -
                     leftRatio -
                     rightRatio)),
                Math.Max(
                    1f,
                    safeHeight *
                    (1f -
                     topRatio -
                     bottomRatio)));
        }

        private static void NormalizeCropPair(
            ref int leading,
            ref int trailing)
        {
            int total = leading + trailing;

            if (total < 100000)
                return;

            if (total <= 0)
            {
                leading = 0;
                trailing = 0;
                return;
            }

            double scale =
                99999d /
                total;

            leading =
                Math.Max(
                    0,
                    Math.Min(
                        99999,
                        (int)Math.Round(
                            leading *
                            scale)));

            trailing =
                Math.Max(
                    0,
                    99999 -
                    leading);
        }

        public static StringAlignment ToStringAlignment(
            PresentationTextAlignment alignment)
        {
            if (alignment ==
                PresentationTextAlignment.Center)
            {
                return StringAlignment.Center;
            }

            if (alignment ==
                PresentationTextAlignment.Right)
            {
                return StringAlignment.Far;
            }

            return StringAlignment.Near;
        }

        public static GraphicsPath CreateBasicShapePath(
            PresentationShapeKind kind,
            RectangleF rect)
        {
            if (kind == PresentationShapeKind.Ellipse)
                return CreatePresetShapePath(
                    "ellipse",
                    rect);

            if (kind == PresentationShapeKind.Triangle)
                return CreatePresetShapePath(
                    "triangle",
                    rect);

            if (kind == PresentationShapeKind.Diamond)
                return CreatePresetShapePath(
                    "diamond",
                    rect);

            if (kind ==
                PresentationShapeKind.RoundedRectangle)
            {
                return CreatePresetShapePath(
                    "roundRect",
                    rect);
            }

            return CreatePresetShapePath(
                "rect",
                rect);
        }

        public static GraphicsPath CreatePresetShapePath(
            string preset,
            RectangleF rect)
        {
            GraphicsPath path =
                new GraphicsPath();
            string normalized =
                string.IsNullOrEmpty(preset)
                    ? "rect"
                    : preset;

            if (normalized == "rect")
            {
                path.AddRectangle(rect);
                return path;
            }

            if (normalized == "ellipse")
            {
                path.AddEllipse(rect);
                return path;
            }

            if (normalized == "roundRect")
            {
                float radius =
                    Math.Max(
                        3f,
                        Math.Min(
                            rect.Width,
                            rect.Height) *
                        0.12f);
                float diameter =
                    radius * 2f;

                path.AddArc(
                    rect.Left,
                    rect.Top,
                    diameter,
                    diameter,
                    180f,
                    90f);
                path.AddArc(
                    rect.Right - diameter,
                    rect.Top,
                    diameter,
                    diameter,
                    270f,
                    90f);
                path.AddArc(
                    rect.Right - diameter,
                    rect.Bottom - diameter,
                    diameter,
                    diameter,
                    0f,
                    90f);
                path.AddArc(
                    rect.Left,
                    rect.Bottom - diameter,
                    diameter,
                    diameter,
                    90f,
                    90f);
                path.CloseFigure();
                return path;
            }

            if (normalized == "triangle")
            {
                path.AddPolygon(
                    new PointF[]
                    {
                        new PointF(
                            rect.Left +
                            rect.Width / 2f,
                            rect.Top),
                        new PointF(
                            rect.Right,
                            rect.Bottom),
                        new PointF(
                            rect.Left,
                            rect.Bottom)
                    });
                return path;
            }

            if (normalized == "diamond")
            {
                path.AddPolygon(
                    new PointF[]
                    {
                        new PointF(
                            rect.Left +
                            rect.Width / 2f,
                            rect.Top),
                        new PointF(
                            rect.Right,
                            rect.Top +
                            rect.Height / 2f),
                        new PointF(
                            rect.Left +
                            rect.Width / 2f,
                            rect.Bottom),
                        new PointF(
                            rect.Left,
                            rect.Top +
                            rect.Height / 2f)
                    });
                return path;
            }

            path.AddRectangle(rect);
            return path;
        }
    }

    internal sealed class PresentationTextRenderFragment
    {
        public string Text;
        public string FontFamily;
        public float FontSize;
        public FontStyle FontStyle;
        public Color Color;
        public float Width;
        public float Height;
    }

    internal sealed class PresentationTextRenderLine
    {
        public readonly List<PresentationTextRenderFragment> Fragments =
            new List<PresentationTextRenderFragment>();
        public float Width;
        public float Height;
        public float SpaceBefore;
        public float SpaceAfter;
        public PresentationTextAlignment Alignment;
    }

    internal static class PresentationModelTextRenderer
    {
        public static void DrawTextBox(
            Graphics graphics,
            PresentationTextBox box,
            RectangleF rect,
            float scale,
            Color fallbackColor)
        {
            if (graphics == null ||
                box == null ||
                rect.Width <= 0f ||
                rect.Height <= 0f)
            {
                return;
            }

            float safeScale =
                Math.Max(
                    0.1f,
                    scale);
            List<PresentationTextRenderLine> lines =
                BuildLines(
                    graphics,
                    box,
                    Math.Max(
                        1f,
                        rect.Width),
                    safeScale,
                    fallbackColor);

            if (lines.Count == 0)
                return;

            float totalHeight = 0f;

            for (int i = 0;
                 i < lines.Count;
                 i++)
            {
                PresentationTextRenderLine line =
                    lines[i];

                totalHeight +=
                    line.SpaceBefore +
                    Math.Max(1f, line.Height) +
                    line.SpaceAfter;
            }

            float y =
                rect.Top +
                Math.Max(
                    0f,
                    (rect.Height -
                     totalHeight) /
                    2f);

            GraphicsState state =
                graphics.Save();

            try
            {
                graphics.SetClip(
                    rect,
                    CombineMode.Intersect);

                using (StringFormat format =
                    (StringFormat)
                    StringFormat.GenericTypographic
                        .Clone())
                {
                    format.FormatFlags |=
                        StringFormatFlags
                            .MeasureTrailingSpaces;

                    for (int i = 0;
                         i < lines.Count;
                         i++)
                    {
                        PresentationTextRenderLine line =
                            lines[i];
                        y += line.SpaceBefore;

                        float x =
                            CalculateLineX(
                                line,
                                rect);

                        for (int f = 0;
                             f < line.Fragments.Count;
                             f++)
                        {
                            PresentationTextRenderFragment fragment =
                                line.Fragments[f];

                            if (fragment == null ||
                                string.IsNullOrEmpty(
                                    fragment.Text))
                            {
                                continue;
                            }

                            using (Font font =
                                PresentationRenderPrimitives
                                    .SafeFont(
                                        fragment.FontFamily,
                                        fragment.FontSize,
                                        fragment.FontStyle))
                            using (Brush brush =
                                new SolidBrush(
                                    fragment.Color))
                            {
                                float top =
                                    y +
                                    Math.Max(
                                        0f,
                                        (line.Height -
                                         fragment.Height) /
                                        2f);

                                graphics.DrawString(
                                    fragment.Text,
                                    font,
                                    brush,
                                    new PointF(
                                        x,
                                        top),
                                    format);
                            }

                            x += fragment.Width;
                        }

                        y +=
                            Math.Max(
                                1f,
                                line.Height) +
                            line.SpaceAfter;

                        if (y > rect.Bottom)
                            break;
                    }
                }
            }
            finally
            {
                graphics.Restore(state);
            }
        }

        private static List<PresentationTextRenderLine>
            BuildLines(
                Graphics graphics,
                PresentationTextBox box,
                float maxWidth,
                float scale,
                Color fallbackColor)
        {
            List<PresentationTextRenderLine> lines =
                new List<PresentationTextRenderLine>();

            if (box.HasRichText)
            {
                for (int p = 0;
                     p < box.RichParagraphs.Count;
                     p++)
                {
                    PresentationTextParagraph paragraph =
                        box.RichParagraphs[p];

                    if (paragraph == null)
                        continue;

                    AppendParagraph(
                        graphics,
                        lines,
                        paragraph,
                        box,
                        maxWidth,
                        scale,
                        fallbackColor);
                }
            }
            else
            {
                PresentationTextParagraph paragraph =
                    new PresentationTextParagraph();
                paragraph.Alignment =
                    box.Alignment;

                PresentationTextRun run =
                    CreateDefaultRun(box);
                run.Text =
                    NormalizeText(
                        box.Text);
                paragraph.Runs.Add(run);

                AppendParagraph(
                    graphics,
                    lines,
                    paragraph,
                    box,
                    maxWidth,
                    scale,
                    fallbackColor);
            }

            return lines;
        }

        private static void AppendParagraph(
            Graphics graphics,
            List<PresentationTextRenderLine> output,
            PresentationTextParagraph paragraph,
            PresentationTextBox box,
            float maxWidth,
            float scale,
            Color fallbackColor)
        {
            PresentationTextRenderLine line =
                NewLine(
                    paragraph.Alignment,
                    paragraph.SpaceBeforePoints *
                    scale);

            if (!string.IsNullOrEmpty(
                    paragraph.BulletText))
            {
                PresentationTextRun bullet =
                    paragraph.Runs.Count > 0 &&
                    paragraph.Runs[0] != null
                        ? paragraph.Runs[0].Clone()
                        : CreateDefaultRun(box);
                bullet.Text =
                    paragraph.BulletText +
                    " ";
                AppendRun(
                    graphics,
                    output,
                    ref line,
                    bullet,
                    box,
                    maxWidth,
                    scale,
                    fallbackColor);
            }

            for (int r = 0;
                 r < paragraph.Runs.Count;
                 r++)
            {
                PresentationTextRun run =
                    paragraph.Runs[r];

                if (run == null)
                    continue;

                AppendRun(
                    graphics,
                    output,
                    ref line,
                    run,
                    box,
                    maxWidth,
                    scale,
                    fallbackColor);
            }

            if (line.Fragments.Count > 0 ||
                output.Count == 0)
            {
                line.SpaceAfter =
                    paragraph.SpaceAfterPoints *
                    scale;
                output.Add(line);
            }
            else if (output.Count > 0)
            {
                output[
                    output.Count - 1]
                    .SpaceAfter +=
                    paragraph.SpaceAfterPoints *
                    scale;
            }
        }

        private static void AppendRun(
            Graphics graphics,
            List<PresentationTextRenderLine> output,
            ref PresentationTextRenderLine line,
            PresentationTextRun run,
            PresentationTextBox box,
            float maxWidth,
            float scale,
            Color fallbackColor)
        {
            string text =
                NormalizeText(
                    run.Text);

            if (text.Length == 0)
                return;

            List<string> tokens =
                Tokenize(text);

            for (int i = 0;
                 i < tokens.Count;
                 i++)
            {
                string token =
                    tokens[i];

                if (token == "\n")
                {
                    output.Add(line);
                    line =
                        NewLine(
                            line.Alignment,
                            0f);
                    continue;
                }

                PresentationTextRenderFragment fragment =
                    CreateFragment(
                        graphics,
                        token,
                        run,
                        box,
                        scale,
                        fallbackColor);

                bool whitespaceOnly =
                    string.IsNullOrWhiteSpace(
                        token);

                if (!whitespaceOnly &&
                    line.Fragments.Count > 0 &&
                    line.Width +
                    fragment.Width >
                    maxWidth)
                {
                    output.Add(line);
                    line =
                        NewLine(
                            line.Alignment,
                            0f);
                }

                line.Fragments.Add(fragment);
                line.Width +=
                    fragment.Width;
                line.Height =
                    Math.Max(
                        line.Height,
                        fragment.Height);
            }
        }

        private static PresentationTextRenderFragment
            CreateFragment(
                Graphics graphics,
                string text,
                PresentationTextRun run,
                PresentationTextBox box,
                float scale,
                Color fallbackColor)
        {
            string family =
                string.IsNullOrEmpty(
                    run.FontFamily)
                    ? box.FontFamily
                    : run.FontFamily;
            float points =
                run.FontSizePoints > 0f
                    ? run.FontSizePoints
                    : box.FontSizePoints;
            float displaySize =
                Math.Max(
                    6f,
                    points *
                    scale);
            FontStyle style =
                FontStyle.Regular;

            if (run.Bold)
                style |= FontStyle.Bold;
            if (run.Italic)
                style |= FontStyle.Italic;
            if (run.Underline)
                style |= FontStyle.Underline;

            PresentationTextRenderFragment fragment =
                new PresentationTextRenderFragment();
            fragment.Text = text;
            fragment.FontFamily = family;
            fragment.FontSize = displaySize;
            fragment.FontStyle = style;
            fragment.Color =
                PresentationRenderPrimitives
                    .ParseHexColor(
                        string.IsNullOrEmpty(
                            run.ColorHex)
                            ? box.ColorHex
                            : run.ColorHex,
                        fallbackColor);

            using (Font font =
                PresentationRenderPrimitives
                    .SafeFont(
                        family,
                        displaySize,
                        style))
            using (StringFormat format =
                (StringFormat)
                StringFormat.GenericTypographic
                    .Clone())
            {
                format.FormatFlags |=
                    StringFormatFlags
                        .MeasureTrailingSpaces;

                SizeF measured =
                    graphics.MeasureString(
                        text,
                        font,
                        int.MaxValue,
                        format);
                fragment.Width =
                    Math.Max(
                        0f,
                        measured.Width);
                fragment.Height =
                    Math.Max(
                        font.Height,
                        measured.Height);
            }

            return fragment;
        }

        private static PresentationTextRenderLine NewLine(
            PresentationTextAlignment alignment,
            float before)
        {
            PresentationTextRenderLine line =
                new PresentationTextRenderLine();
            line.Alignment = alignment;
            line.SpaceBefore =
                Math.Max(
                    0f,
                    before);
            return line;
        }

        private static float CalculateLineX(
            PresentationTextRenderLine line,
            RectangleF rect)
        {
            if (line.Alignment ==
                PresentationTextAlignment.Center)
            {
                return
                    rect.Left +
                    Math.Max(
                        0f,
                        (rect.Width -
                         line.Width) /
                        2f);
            }

            if (line.Alignment ==
                PresentationTextAlignment.Right)
            {
                return
                    rect.Right -
                    Math.Min(
                        rect.Width,
                        line.Width);
            }

            return rect.Left;
        }

        private static PresentationTextRun
            CreateDefaultRun(
                PresentationTextBox box)
        {
            PresentationTextRun run =
                new PresentationTextRun();
            run.FontFamily =
                box.FontFamily;
            run.FontSizePoints =
                box.FontSizePoints;
            run.Bold =
                box.Bold;
            run.Italic =
                box.Italic;
            run.ColorHex =
                box.ColorHex;
            return run;
        }

        private static List<string> Tokenize(
            string value)
        {
            List<string> tokens =
                new List<string>();
            StringBuilder current =
                new StringBuilder();
            bool? whitespace = null;

            for (int i = 0;
                 i < value.Length;
                 i++)
            {
                char ch =
                    value[i];

                if (ch == '\n')
                {
                    FlushToken(
                        tokens,
                        current);
                    tokens.Add("\n");
                    whitespace = null;
                    continue;
                }

                bool isWhitespace =
                    char.IsWhiteSpace(ch);

                if (whitespace.HasValue &&
                    whitespace.Value !=
                    isWhitespace)
                {
                    FlushToken(
                        tokens,
                        current);
                }

                current.Append(ch);
                whitespace =
                    isWhitespace;
            }

            FlushToken(
                tokens,
                current);
            return tokens;
        }

        private static void FlushToken(
            List<string> tokens,
            StringBuilder current)
        {
            if (current.Length == 0)
                return;

            tokens.Add(
                current.ToString());
            current.Length = 0;
        }

        private static string NormalizeText(
            string value)
        {
            return (value ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");
        }
    }

    internal static class PresentationRenderPrimitivesDiagnostics
    {
        public static void Validate()
        {
            Color parsed =
                PresentationRenderPrimitives
                    .ParseHexColor(
                        "#336699",
                        Color.Magenta);

            if (parsed.R != 0x33 ||
                parsed.G != 0x66 ||
                parsed.B != 0x99)
            {
                throw new InvalidOperationException(
                    "Shared render color parsing failed.");
            }

            RectangleF bounds =
                new RectangleF(
                    10f,
                    20f,
                    140f,
                    80f);

            using (GraphicsPath rounded =
                PresentationRenderPrimitives
                    .CreatePresetShapePath(
                        "roundRect",
                        bounds))
            using (GraphicsPath diamond =
                PresentationRenderPrimitives
                    .CreateBasicShapePath(
                        PresentationShapeKind.Diamond,
                        bounds))
            {
                RectangleF roundedBounds =
                    rounded.GetBounds();
                RectangleF diamondBounds =
                    diamond.GetBounds();

                if (roundedBounds.Width < 130f ||
                    roundedBounds.Height < 70f ||
                    diamondBounds.Width < 130f ||
                    diamondBounds.Height < 70f)
                {
                    throw new InvalidOperationException(
                        "Shared render geometry lost the requested bounds.");
                }
            }

            PresentationTextBox box =
                new PresentationTextBox();
            box.Text =
                "Shared rich text render";
            int start =
                box.Text.IndexOf(
                    "rich",
                    StringComparison.Ordinal);
            RichTextFormatChange color =
                new RichTextFormatChange();
            color.Field =
                RichTextFormatField.Color;
            color.ColorHex =
                "C02020";
            RichTextSelectionEditor.ApplySelection(
                box,
                start,
                4,
                color);

            using (Bitmap bitmap =
                new Bitmap(
                    480,
                    240))
            using (Graphics graphics =
                Graphics.FromImage(
                    bitmap))
            {
                graphics.Clear(
                    Color.White);
                PresentationModelTextRenderer.DrawTextBox(
                    graphics,
                    box,
                    new RectangleF(
                        20f,
                        20f,
                        440f,
                        200f),
                    1f,
                    Color.Black);

                bool foundInk = false;

                for (int y = 20;
                     y < 220 &&
                     !foundInk;
                     y += 4)
                {
                    for (int x = 20;
                         x < 460;
                         x += 4)
                    {
                        Color pixel =
                            bitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R < 245 ||
                            pixel.G < 245 ||
                            pixel.B < 245)
                        {
                            foundInk = true;
                            break;
                        }
                    }
                }

                if (!foundInk)
                {
                    throw new InvalidOperationException(
                        "Shared model text renderer produced no visible output.");
                }
            }
        }
    }
}
