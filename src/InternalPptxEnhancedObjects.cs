using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Xml;

namespace PptxViewer
{
    internal static partial class InternalPptxRenderer
    {
        private sealed class SmartNode
        {
            public string Id;
            public string Label;
            public readonly List<string> Children = new List<string>();
            public int Depth;
        }

        private sealed class SvgGradientStop
        {
            public float Offset;
            public Color Color;
        }

        private static void DrawEnhancedShape(
            ZipArchive zip,
            Graphics g,
            XmlNode shape,
            Dictionary<string, string> rels,
            TransformContext ctx,
            Dictionary<string, Color> theme,
            Dictionary<string, RectangleF> placeholderRects,
            int slideNumber)
        {
            XmlNode spPr = DirectChild(shape, "spPr");
            XmlNode prstGeom = spPr != null ? FindFirst(spPr, "prstGeom") : null;
            string preset = prstGeom != null ? GetAttr(prstGeom, "prst") : string.Empty;

            if (!IsEnhancedPreset(preset))
            {
                DrawShape(zip, g, shape, rels, ctx, theme, placeholderRects, slideNumber);
                return;
            }

            RectangleF rect;
            if (!TryGetRect(shape, ctx, out rect) &&
                !TryGetPlaceholderRect(shape, placeholderRects, out rect))
            {
                return;
            }

            bool noFill = spPr != null && FindFirst(spPr, "noFill") != null;
            Brush fillBrush = !noFill && spPr != null
                ? CreateFillBrush(spPr, rect, theme)
                : null;
            Color? line = spPr != null ? ReadLineColor(spPr, theme) : null;
            GraphicsState state = g.Save();

            try
            {
                ApplyRotation(g, shape, rect);

                using (GraphicsPath path = BuildEnhancedPresetPath(preset, rect))
                {
                    DrawShapeVisualEffects(g, spPr, path, rect, theme);
                    bool pictureFill = DrawShapePictureFill(zip, g, spPr, rels, path, rect);

                    if (!pictureFill && fillBrush != null)
                        g.FillPath(fillBrush, path);

                    DrawShapePostFillEffects(
                        g,
                        spPr,
                        path,
                        rect,
                        theme);

                    if (line.HasValue)
                    {
                        using (Pen pen = CreateLinePen(spPr, theme, line.Value))
                            g.DrawPath(pen, path);
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

        private static bool IsEnhancedPreset(string preset)
        {
            return preset == "pentagon" ||
                   preset == "octagon" ||
                   preset == "star5" ||
                   preset == "star6" ||
                   preset == "plus" ||
                   preset == "chevron" ||
                   preset == "homePlate" ||
                   preset == "trapezoid" ||
                   preset == "leftArrow" ||
                   preset == "rightArrow" ||
                   preset == "upArrow" ||
                   preset == "downArrow";
        }

        private static GraphicsPath BuildEnhancedPresetPath(
            string preset,
            RectangleF r)
        {
            GraphicsPath path = new GraphicsPath();
            PointF[] points = null;

            if (preset == "pentagon")
                points = RegularPolygon(r, 5, -90f);
            else if (preset == "octagon")
                points = RegularPolygon(r, 8, 22.5f);
            else if (preset == "star5")
                points = StarPolygon(r, 5, 0.43f, -90f);
            else if (preset == "star6")
                points = StarPolygon(r, 6, 0.48f, -90f);
            else if (preset == "plus")
            {
                float x1 = r.Left + r.Width * 0.34f;
                float x2 = r.Left + r.Width * 0.66f;
                float y1 = r.Top + r.Height * 0.34f;
                float y2 = r.Top + r.Height * 0.66f;
                points = new PointF[]
                {
                    new PointF(x1, r.Top), new PointF(x2, r.Top),
                    new PointF(x2, y1), new PointF(r.Right, y1),
                    new PointF(r.Right, y2), new PointF(x2, y2),
                    new PointF(x2, r.Bottom), new PointF(x1, r.Bottom),
                    new PointF(x1, y2), new PointF(r.Left, y2),
                    new PointF(r.Left, y1), new PointF(x1, y1)
                };
            }
            else if (preset == "chevron")
            {
                float dx = r.Width * 0.28f;
                points = new PointF[]
                {
                    new PointF(r.Left, r.Top),
                    new PointF(r.Right - dx, r.Top),
                    new PointF(r.Right, r.Top + r.Height / 2f),
                    new PointF(r.Right - dx, r.Bottom),
                    new PointF(r.Left, r.Bottom),
                    new PointF(r.Left + dx, r.Top + r.Height / 2f)
                };
            }
            else if (preset == "homePlate")
            {
                float dx = r.Width * 0.22f;
                points = new PointF[]
                {
                    new PointF(r.Left, r.Top),
                    new PointF(r.Right - dx, r.Top),
                    new PointF(r.Right, r.Top + r.Height / 2f),
                    new PointF(r.Right - dx, r.Bottom),
                    new PointF(r.Left, r.Bottom)
                };
            }
            else if (preset == "trapezoid")
            {
                float dx = r.Width * 0.18f;
                points = new PointF[]
                {
                    new PointF(r.Left + dx, r.Top),
                    new PointF(r.Right - dx, r.Top),
                    new PointF(r.Right, r.Bottom),
                    new PointF(r.Left, r.Bottom)
                };
            }
            else if (preset == "leftArrow")
                points = ArrowPolygon(r, false, true);
            else if (preset == "rightArrow")
                points = ArrowPolygon(r, false, false);
            else if (preset == "upArrow")
                points = ArrowPolygon(r, true, true);
            else if (preset == "downArrow")
                points = ArrowPolygon(r, true, false);

            if (points == null || points.Length < 3)
                path.AddRectangle(r);
            else
            {
                path.AddPolygon(points);
                path.CloseFigure();
            }

            return path;
        }

        private static PointF[] RegularPolygon(
            RectangleF r,
            int count,
            float startDegrees)
        {
            PointF[] points = new PointF[count];
            float cx = r.Left + r.Width / 2f;
            float cy = r.Top + r.Height / 2f;
            float rx = r.Width / 2f;
            float ry = r.Height / 2f;

            for (int i = 0; i < count; i++)
            {
                double angle = (startDegrees + i * 360.0 / count) * Math.PI / 180.0;
                points[i] = new PointF(
                    cx + (float)Math.Cos(angle) * rx,
                    cy + (float)Math.Sin(angle) * ry);
            }
            return points;
        }

        private static PointF[] StarPolygon(
            RectangleF r,
            int pointsCount,
            float innerRatio,
            float startDegrees)
        {
            PointF[] points = new PointF[pointsCount * 2];
            float cx = r.Left + r.Width / 2f;
            float cy = r.Top + r.Height / 2f;
            float rx = r.Width / 2f;
            float ry = r.Height / 2f;

            for (int i = 0; i < points.Length; i++)
            {
                bool outer = i % 2 == 0;
                float ratio = outer ? 1f : innerRatio;
                double angle = (startDegrees + i * 180.0 / pointsCount) * Math.PI / 180.0;
                points[i] = new PointF(
                    cx + (float)Math.Cos(angle) * rx * ratio,
                    cy + (float)Math.Sin(angle) * ry * ratio);
            }
            return points;
        }

        private static PointF[] ArrowPolygon(
            RectangleF r,
            bool vertical,
            bool reverse)
        {
            if (!vertical)
            {
                float shaftTop = r.Top + r.Height * 0.30f;
                float shaftBottom = r.Bottom - r.Height * 0.30f;
                float head = r.Width * 0.38f;
                PointF[] right = new PointF[]
                {
                    new PointF(r.Left, shaftTop),
                    new PointF(r.Right - head, shaftTop),
                    new PointF(r.Right - head, r.Top),
                    new PointF(r.Right, r.Top + r.Height / 2f),
                    new PointF(r.Right - head, r.Bottom),
                    new PointF(r.Right - head, shaftBottom),
                    new PointF(r.Left, shaftBottom)
                };
                if (!reverse) return right;
                return MirrorHorizontal(right, r);
            }

            float shaftLeft = r.Left + r.Width * 0.30f;
            float shaftRight = r.Right - r.Width * 0.30f;
            float headY = r.Height * 0.38f;
            PointF[] down = new PointF[]
            {
                new PointF(shaftLeft, r.Top),
                new PointF(shaftRight, r.Top),
                new PointF(shaftRight, r.Bottom - headY),
                new PointF(r.Right, r.Bottom - headY),
                new PointF(r.Left + r.Width / 2f, r.Bottom),
                new PointF(r.Left, r.Bottom - headY),
                new PointF(shaftLeft, r.Bottom - headY)
            };
            if (!reverse) return down;
            return MirrorVertical(down, r);
        }

        private static PointF[] MirrorHorizontal(PointF[] points, RectangleF r)
        {
            PointF[] result = new PointF[points.Length];
            for (int i = 0; i < points.Length; i++)
                result[i] = new PointF(r.Left + r.Right - points[i].X, points[i].Y);
            return result;
        }

        private static PointF[] MirrorVertical(PointF[] points, RectangleF r)
        {
            PointF[] result = new PointF[points.Length];
            for (int i = 0; i < points.Length; i++)
                result[i] = new PointF(points[i].X, r.Top + r.Bottom - points[i].Y);
            return result;
        }

        private static void DrawEnhancedPicture(
            ZipArchive zip,
            Graphics g,
            XmlNode pic,
            Dictionary<string, string> rels,
            TransformContext ctx,
            Dictionary<string, Color> theme)
        {
            RectangleF rect;
            if (!TryGetRect(pic, ctx, out rect))
                return;

            XmlNode blip = FindFirst(pic, "blip");
            string rid = GetRelationshipId(blip);
            if (string.IsNullOrEmpty(rid) || rels == null || !rels.ContainsKey(rid))
            {
                DrawPlaceholder(g, rect, "Image");
                return;
            }

            ZipArchiveEntry entry = zip.GetEntry(rels[rid]);
            if (entry == null)
            {
                DrawPlaceholder(g, rect, "Image");
                return;
            }

            try
            {
                using (Stream stream = entry.Open())
                using (Image sourceImage = Image.FromStream(stream))
                using (Bitmap image = new Bitmap(sourceImage))
                {
                    RectangleF source = ReadPictureCrop(pic, image.Width, image.Height);
                    GraphicsState state = g.Save();
                    try
                    {
                        ApplyRotation(g, pic, rect);
                        DrawImageWithDrawingEffects(
                            g,
                            image,
                            source,
                            rect,
                            blip,
                            theme);
                    }
                    finally
                    {
                        g.Restore(state);
                    }
                }
            }
            catch
            {
                if (!TryDrawEnhancedSvg(entry, g, rect))
                    DrawPicture(zip, g, pic, rels, ctx);
            }
        }

        private static RectangleF ReadPictureCrop(
            XmlNode pic,
            int width,
            int height)
        {
            RectangleF source = new RectangleF(0, 0, width, height);
            XmlNode srcRect = FindFirst(pic, "srcRect");
            if (srcRect == null)
                return source;

            float left = Clamp01(GetLong(srcRect, "l", 0) / 100000f);
            float top = Clamp01(GetLong(srcRect, "t", 0) / 100000f);
            float right = Clamp01(GetLong(srcRect, "r", 0) / 100000f);
            float bottom = Clamp01(GetLong(srcRect, "b", 0) / 100000f);

            source.X = width * left;
            source.Y = height * top;
            source.Width = Math.Max(1f, width * (1f - left - right));
            source.Height = Math.Max(1f, height * (1f - top - bottom));
            return source;
        }

        private static float Clamp01(float value)
        {
            return Math.Max(0f, Math.Min(0.99f, value));
        }

        private static void DrawImageWithDrawingEffects(
            Graphics g,
            Bitmap image,
            RectangleF source,
            RectangleF target,
            XmlNode blip,
            Dictionary<string, Color> theme)
        {
            float alpha = 1f;

            XmlNode alphaNode =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "alphaModFix");

            if (alphaNode != null)
            {
                alpha =
                    Math.Max(
                        0f,
                        Math.Min(
                            1f,
                            GetLong(
                                alphaNode,
                                "amt",
                                100000) /
                            100000f));
            }

            bool grayscale =
                blip != null &&
                FindFirst(
                    blip,
                    "grayscl") != null;

            XmlNode biLevel =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "biLevel");

            XmlNode lum =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "lum");

            XmlNode duotone =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "duotone");

            XmlNode colorChange =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "clrChange");

            XmlNode blur =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "blur");

            Bitmap blurredImage =
                CreateBlurredImageApproximation(
                    image,
                    blur);

            Bitmap renderImage =
                blurredImage ??
                image;

            XmlNode artisticBlur =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "artisticBlur");

            Bitmap artisticBlurImage =
                CreateArtisticBlurImageApproximation(
                    renderImage,
                    artisticBlur);

            if (artisticBlurImage != null)
            {
                renderImage =
                    artisticBlurImage;
            }

            XmlNode sharpenSoften =
                blip == null
                    ? null
                    : FindFirst(
                        blip,
                        "sharpenSoften");

            Bitmap sharpenSoftenImage =
                CreateSharpenSoftenImageApproximation(
                    renderImage,
                    sharpenSoften);

            if (sharpenSoftenImage != null)
            {
                renderImage =
                    sharpenSoftenImage;
            }

            try
            {
            float brightness =
                lum == null
                    ? 0f
                    : Math.Max(
                        -1f,
                        Math.Min(
                            1f,
                            GetLong(
                                lum,
                                "bright",
                                0) /
                            100000f));

            float contrast =
                lum == null
                    ? 0f
                    : Math.Max(
                        -1f,
                        Math.Min(
                            1f,
                            GetLong(
                                lum,
                                "contrast",
                                0) /
                            100000f));

            Color duoLow;
            Color duoHigh;
            bool hasDuotone =
                TryReadDuotoneColors(
                    duotone,
                    theme,
                    out duoLow,
                    out duoHigh);

            Color changeFrom;
            Color changeTo;
            bool hasColorChange =
                TryReadColorChange(
                    colorChange,
                    theme,
                    out changeFrom,
                    out changeTo);

            bool needsAttributes =
                grayscale ||
                biLevel != null ||
                hasDuotone ||
                hasColorChange ||
                alpha < 0.999f ||
                Math.Abs(brightness) >
                    0.001f ||
                Math.Abs(contrast) >
                    0.001f;

            if (!needsAttributes)
            {
                g.DrawImage(
                    renderImage,
                    Rectangle.Round(target),
                    source.X,
                    source.Y,
                    source.Width,
                    source.Height,
                    GraphicsUnit.Pixel);
                return;
            }

            using (ImageAttributes attributes =
                new ImageAttributes())
            {
                ColorMatrix matrix;

                if (hasDuotone)
                {
                    float scale =
                        Math.Max(
                            0f,
                            1f +
                            contrast);
                    float offset =
                        brightness +
                        (1f -
                         scale) *
                        0.5f;

                    float lowR =
                        duoLow.R /
                        255f;
                    float lowG =
                        duoLow.G /
                        255f;
                    float lowB =
                        duoLow.B /
                        255f;

                    float deltaR =
                        (duoHigh.R -
                         duoLow.R) /
                        255f;
                    float deltaG =
                        (duoHigh.G -
                         duoLow.G) /
                        255f;
                    float deltaB =
                        (duoHigh.B -
                         duoLow.B) /
                        255f;

                    matrix =
                        new ColorMatrix(
                            new float[][]
                            {
                                new float[]
                                {
                                    0.299f *
                                        scale *
                                        deltaR,
                                    0.299f *
                                        scale *
                                        deltaG,
                                    0.299f *
                                        scale *
                                        deltaB,
                                    0f,
                                    0f
                                },
                                new float[]
                                {
                                    0.587f *
                                        scale *
                                        deltaR,
                                    0.587f *
                                        scale *
                                        deltaG,
                                    0.587f *
                                        scale *
                                        deltaB,
                                    0f,
                                    0f
                                },
                                new float[]
                                {
                                    0.114f *
                                        scale *
                                        deltaR,
                                    0.114f *
                                        scale *
                                        deltaG,
                                    0.114f *
                                        scale *
                                        deltaB,
                                    0f,
                                    0f
                                },
                                new float[]
                                {
                                    0f,
                                    0f,
                                    0f,
                                    alpha,
                                    0f
                                },
                                new float[]
                                {
                                    Math.Max(
                                        -1f,
                                        Math.Min(
                                            1f,
                                            lowR +
                                            offset *
                                            deltaR)),
                                    Math.Max(
                                        -1f,
                                        Math.Min(
                                            1f,
                                            lowG +
                                            offset *
                                            deltaG)),
                                    Math.Max(
                                        -1f,
                                        Math.Min(
                                            1f,
                                            lowB +
                                            offset *
                                            deltaB)),
                                    0f,
                                    1f
                                }
                            });
                }
                else if (grayscale)
                {
                    matrix =
                        new ColorMatrix(
                            new float[][]
                            {
                                new float[]
                                {
                                    0.299f,
                                    0.299f,
                                    0.299f,
                                    0f,
                                    0f
                                },
                                new float[]
                                {
                                    0.587f,
                                    0.587f,
                                    0.587f,
                                    0f,
                                    0f
                                },
                                new float[]
                                {
                                    0.114f,
                                    0.114f,
                                    0.114f,
                                    0f,
                                    0f
                                },
                                new float[]
                                {
                                    0f,
                                    0f,
                                    0f,
                                    alpha,
                                    0f
                                },
                                new float[]
                                {
                                    brightness,
                                    brightness,
                                    brightness,
                                    0f,
                                    1f
                                }
                            });
                }
                else
                {
                    matrix =
                        new ColorMatrix();

                    float scale =
                        Math.Max(
                            0f,
                            1f +
                            contrast);
                    float offset =
                        brightness +
                        (1f -
                         scale) *
                        0.5f;

                    matrix.Matrix00 =
                        scale;
                    matrix.Matrix11 =
                        scale;
                    matrix.Matrix22 =
                        scale;
                    matrix.Matrix33 =
                        alpha;
                    matrix.Matrix40 =
                        offset;
                    matrix.Matrix41 =
                        offset;
                    matrix.Matrix42 =
                        offset;
                }

                attributes.SetColorMatrix(
                    matrix,
                    ColorMatrixFlag.Default,
                    ColorAdjustType.Bitmap);

                if (hasColorChange)
                {
                    ColorMap map =
                        new ColorMap();

                    map.OldColor =
                        changeFrom;
                    map.NewColor =
                        changeTo;

                    attributes.SetRemapTable(
                        new ColorMap[]
                        {
                            map
                        },
                        ColorAdjustType.Bitmap);
                }

                if (biLevel != null)
                {
                    float threshold =
                        Math.Max(
                            0f,
                            Math.Min(
                                1f,
                                GetLong(
                                    biLevel,
                                    "thresh",
                                    50000) /
                                100000f));

                    attributes.SetThreshold(
                        threshold,
                        ColorAdjustType.Bitmap);
                }

                g.DrawImage(
                    renderImage,
                    Rectangle.Round(target),
                    source.X,
                    source.Y,
                    source.Width,
                    source.Height,
                    GraphicsUnit.Pixel,
                    attributes);
            }
            }
            finally
            {
                if (sharpenSoftenImage != null)
                {
                    sharpenSoftenImage.Dispose();
                }

                if (artisticBlurImage != null)
                {
                    artisticBlurImage.Dispose();
                }

                if (blurredImage != null)
                {
                    blurredImage.Dispose();
                }
            }
        }

        private static Bitmap CreateSharpenSoftenImageApproximation(
            Bitmap image,
            XmlNode effect)
        {
            if (image == null ||
                effect == null ||
                image.Width <= 0 ||
                image.Height <= 0)
            {
                return null;
            }

            long rawAmount =
                GetLong(
                    effect,
                    "amount",
                    0);

            int amount =
                (int)Math.Max(
                    -100000L,
                    Math.Min(
                        100000L,
                        rawAmount));

            if (Math.Abs(amount) < 1000)
            {
                return null;
            }

            using (Bitmap normalized =
                new Bitmap(
                    image.Width,
                    image.Height,
                    PixelFormat.Format32bppArgb))
            {
                using (Graphics normalizedGraphics =
                    Graphics.FromImage(
                        normalized))
                {
                    normalizedGraphics.CompositingMode =
                        CompositingMode.SourceCopy;
                    normalizedGraphics.DrawImage(
                        image,
                        new Rectangle(
                            0,
                            0,
                            normalized.Width,
                            normalized.Height),
                        0,
                        0,
                        image.Width,
                        image.Height,
                        GraphicsUnit.Pixel);
                }

                Rectangle bounds =
                    new Rectangle(
                        0,
                        0,
                        normalized.Width,
                        normalized.Height);

                BitmapData sourceData =
                    normalized.LockBits(
                        bounds,
                        ImageLockMode.ReadOnly,
                        PixelFormat.Format32bppArgb);

                byte[] sourceBytes;

                try
                {
                    if (sourceData.Stride <= 0)
                    {
                        return null;
                    }

                    sourceBytes =
                        new byte[
                            sourceData.Stride *
                            normalized.Height];

                    Marshal.Copy(
                        sourceData.Scan0,
                        sourceBytes,
                        0,
                        sourceBytes.Length);
                }
                finally
                {
                    normalized.UnlockBits(
                        sourceData);
                }

                int stride =
                    normalized.Width *
                    4;

                if (sourceBytes.Length <
                    stride *
                    normalized.Height)
                {
                    stride =
                        sourceBytes.Length /
                        normalized.Height;
                }

                byte[] blurredBytes =
                    new byte[
                        sourceBytes.Length];

                Array.Copy(
                    sourceBytes,
                    blurredBytes,
                    sourceBytes.Length);

                for (int y = 0;
                     y < normalized.Height;
                     y++)
                {
                    for (int x = 0;
                         x < normalized.Width;
                         x++)
                    {
                        int pixelOffset =
                            y *
                            stride +
                            x *
                            4;

                        for (int channel = 0;
                             channel < 3;
                             channel++)
                        {
                            int sum = 0;
                            int count = 0;

                            for (int oy = -1;
                                 oy <= 1;
                                 oy++)
                            {
                                int sampleY =
                                    y +
                                    oy;

                                if (sampleY < 0 ||
                                    sampleY >=
                                        normalized.Height)
                                {
                                    continue;
                                }

                                for (int ox = -1;
                                     ox <= 1;
                                     ox++)
                                {
                                    int sampleX =
                                        x +
                                        ox;

                                    if (sampleX < 0 ||
                                        sampleX >=
                                            normalized.Width)
                                    {
                                        continue;
                                    }

                                    int sampleOffset =
                                        sampleY *
                                        stride +
                                        sampleX *
                                        4 +
                                        channel;

                                    sum +=
                                        sourceBytes[
                                            sampleOffset];
                                    count++;
                                }
                            }

                            blurredBytes[
                                pixelOffset +
                                channel] =
                                (byte)(
                                    count > 0
                                        ? sum /
                                          count
                                        : sourceBytes[
                                            pixelOffset +
                                            channel]);
                        }

                        blurredBytes[
                            pixelOffset +
                            3] =
                            sourceBytes[
                                pixelOffset +
                                3];
                    }
                }

                byte[] outputBytes =
                    new byte[
                        sourceBytes.Length];

                Array.Copy(
                    sourceBytes,
                    outputBytes,
                    sourceBytes.Length);

                float strength =
                    Math.Min(
                        1f,
                        Math.Abs(
                            amount) /
                        100000f);

                float sharpenStrength =
                    strength *
                    1.75f;

                for (int y = 0;
                     y < normalized.Height;
                     y++)
                {
                    for (int x = 0;
                         x < normalized.Width;
                         x++)
                    {
                        int pixelOffset =
                            y *
                            stride +
                            x *
                            4;

                        for (int channel = 0;
                             channel < 3;
                             channel++)
                        {
                            float original =
                                sourceBytes[
                                    pixelOffset +
                                    channel];
                            float softened =
                                blurredBytes[
                                    pixelOffset +
                                    channel];

                            float adjusted =
                                amount > 0
                                    ? original +
                                      (original -
                                       softened) *
                                      sharpenStrength
                                    : original +
                                      (softened -
                                       original) *
                                      strength;

                            outputBytes[
                                pixelOffset +
                                channel] =
                                (byte)Math.Max(
                                    0,
                                    Math.Min(
                                        255,
                                        (int)Math.Round(
                                            adjusted)));
                        }

                        outputBytes[
                            pixelOffset +
                            3] =
                            sourceBytes[
                                pixelOffset +
                                3];
                    }
                }

                Bitmap result =
                    new Bitmap(
                        normalized.Width,
                        normalized.Height,
                        PixelFormat.Format32bppArgb);

                BitmapData outputData =
                    result.LockBits(
                        bounds,
                        ImageLockMode.WriteOnly,
                        PixelFormat.Format32bppArgb);

                bool outputUnlocked =
                    false;

                try
                {
                    if (outputData.Stride <= 0 ||
                        outputData.Stride !=
                            stride)
                    {
                        result.UnlockBits(
                            outputData);
                        outputUnlocked =
                            true;
                        result.Dispose();
                        return null;
                    }

                    Marshal.Copy(
                        outputBytes,
                        0,
                        outputData.Scan0,
                        outputBytes.Length);
                }
                finally
                {
                    if (!outputUnlocked)
                    {
                        result.UnlockBits(
                            outputData);
                    }
                }

                return result;
            }
        }

        private static Bitmap CreateBlurredImageApproximation(
            Bitmap image,
            XmlNode blur)
        {
            if (image == null ||
                blur == null)
            {
                return null;
            }

            long radiusEmu =
                Math.Max(
                    0L,
                    GetLong(
                        blur,
                        "rad",
                        0));

            if (radiusEmu <= 0L)
            {
                return null;
            }

            float radiusPoints =
                radiusEmu /
                12700f;

            int radiusPixels =
                Math.Max(
                    1,
                    Math.Min(
                        48,
                        (int)Math.Ceiling(
                            radiusPoints *
                            96f /
                            72f)));

            return CreateBlurredImageApproximation(
                image,
                radiusPixels);
        }

        private static Bitmap CreateArtisticBlurImageApproximation(
            Bitmap image,
            XmlNode effect)
        {
            if (image == null ||
                effect == null)
            {
                return null;
            }

            int radius =
                (int)Math.Max(
                    0L,
                    Math.Min(
                        100L,
                        GetLong(
                            effect,
                            "radius",
                            10)));

            if (radius <= 0)
            {
                return null;
            }

            int radiusPixels =
                Math.Max(
                    1,
                    Math.Min(
                        48,
                        (int)Math.Ceiling(
                            radius *
                            0.48f)));

            return CreateBlurredImageApproximation(
                image,
                radiusPixels);
        }

        private static Bitmap CreateBlurredImageApproximation(
            Bitmap image,
            int radiusPixels)
        {
            if (image == null ||
                radiusPixels <= 0)
            {
                return null;
            }

            radiusPixels =
                Math.Max(
                    1,
                    Math.Min(
                        48,
                        radiusPixels));

            int divisor =
                Math.Max(
                    2,
                    Math.Min(
                        16,
                        1 +
                        radiusPixels /
                        2));

            int reducedWidth =
                Math.Max(
                    1,
                    image.Width /
                    divisor);
            int reducedHeight =
                Math.Max(
                    1,
                    image.Height /
                    divisor);

            using (Bitmap reduced =
                new Bitmap(
                    reducedWidth,
                    reducedHeight,
                    PixelFormat.Format32bppArgb))
            {
                using (Graphics reducedGraphics =
                    Graphics.FromImage(
                        reduced))
                {
                    reducedGraphics.CompositingMode =
                        CompositingMode.SourceCopy;
                    reducedGraphics.CompositingQuality =
                        CompositingQuality.HighQuality;
                    reducedGraphics.InterpolationMode =
                        InterpolationMode.HighQualityBilinear;
                    reducedGraphics.PixelOffsetMode =
                        PixelOffsetMode.HighQuality;
                    reducedGraphics.SmoothingMode =
                        SmoothingMode.HighQuality;

                    reducedGraphics.DrawImage(
                        image,
                        new Rectangle(
                            0,
                            0,
                            reducedWidth,
                            reducedHeight),
                        0,
                        0,
                        image.Width,
                        image.Height,
                        GraphicsUnit.Pixel);
                }

                Bitmap result =
                    new Bitmap(
                        image.Width,
                        image.Height,
                        PixelFormat.Format32bppArgb);

                using (Graphics resultGraphics =
                    Graphics.FromImage(
                        result))
                {
                    resultGraphics.CompositingMode =
                        CompositingMode.SourceCopy;
                    resultGraphics.CompositingQuality =
                        CompositingQuality.HighQuality;
                    resultGraphics.InterpolationMode =
                        InterpolationMode.HighQualityBicubic;
                    resultGraphics.PixelOffsetMode =
                        PixelOffsetMode.HighQuality;
                    resultGraphics.SmoothingMode =
                        SmoothingMode.HighQuality;

                    resultGraphics.DrawImage(
                        reduced,
                        new Rectangle(
                            0,
                            0,
                            result.Width,
                            result.Height),
                        0,
                        0,
                        reduced.Width,
                        reduced.Height,
                        GraphicsUnit.Pixel);
                }

                return result;
            }
        }

        private static bool TryReadDuotoneColors(
            XmlNode duotone,
            Dictionary<string, Color> theme,
            out Color low,
            out Color high)
        {
            low =
                Color.Black;
            high =
                Color.White;

            if (duotone == null)
                return false;

            List<Color> colors =
                new List<Color>();

            for (int i = 0;
                 i < duotone.ChildNodes.Count;
                 i++)
            {
                XmlNode child =
                    duotone.ChildNodes[i];

                if (!IsDrawingColorNode(
                        child))
                {
                    continue;
                }

                Color? color =
                    ReadDrawingEffectColorNode(
                        child,
                        theme);

                if (color.HasValue)
                {
                    colors.Add(
                        color.Value);
                }
            }

            if (colors.Count < 2)
                return false;

            low =
                colors[0];
            high =
                colors[1];
            return true;
        }

        private static bool TryReadColorChange(
            XmlNode colorChange,
            Dictionary<string, Color> theme,
            out Color from,
            out Color to)
        {
            from =
                Color.Empty;
            to =
                Color.Empty;

            if (colorChange == null)
                return false;

            XmlNode fromNode =
                DirectChild(
                    colorChange,
                    "clrFrom");
            XmlNode toNode =
                DirectChild(
                    colorChange,
                    "clrTo");

            Color? fromColor =
                ReadColorFromFill(
                    fromNode,
                    theme);
            Color? toColor =
                ReadColorFromFill(
                    toNode,
                    theme);

            if (!fromColor.HasValue ||
                !toColor.HasValue)
            {
                return false;
            }

            from =
                Color.FromArgb(
                    255,
                    fromColor.Value.R,
                    fromColor.Value.G,
                    fromColor.Value.B);

            to =
                toColor.Value;
            return true;
        }

        private static bool IsDrawingColorNode(
            XmlNode node)
        {
            if (node == null)
                return false;

            string name =
                node.LocalName;

            return name == "srgbClr" ||
                name == "schemeClr" ||
                name == "sysClr" ||
                name == "prstClr";
        }

        private static Color? ReadDrawingEffectColorNode(
            XmlNode colorNode,
            Dictionary<string, Color> theme)
        {
            if (!IsDrawingColorNode(
                    colorNode))
            {
                return null;
            }

            XmlDocument document =
                new XmlDocument();

            XmlElement wrapper =
                document.CreateElement(
                    "wrapper");

            document.AppendChild(
                wrapper);

            wrapper.AppendChild(
                document.ImportNode(
                    colorNode,
                    true));

            return ReadColorFromFill(
                wrapper,
                theme);
        }

        private static bool TryDrawEnhancedSvg(
            ZipArchiveEntry entry,
            Graphics g,
            RectangleF target)
        {
            if (entry == null ||
                !entry.FullName.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
                return false;

            try
            {
                XmlDocument document = new XmlDocument();
                using (Stream input = entry.Open())
                    document.Load(input);

                XmlNode svg = document.DocumentElement;
                if (svg == null || svg.LocalName != "svg")
                    return false;

                float minX = 0f;
                float minY = 0f;
                float width = ParseSvgFloat(GetAttr(svg, "width"), 100f);
                float height = ParseSvgFloat(GetAttr(svg, "height"), 100f);
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
                        width = Math.Max(0.001f, ParseSvgFloat(pieces[2], width));
                        height = Math.Max(0.001f, ParseSvgFloat(pieces[3], height));
                    }
                }

                float sx = target.Width / Math.Max(0.001f, width);
                float sy = target.Height / Math.Max(0.001f, height);
                return DrawEnhancedSvgChildren(g, svg, target, minX, minY, sx, sy);
            }
            catch
            {
                return false;
            }
        }

        private static bool DrawEnhancedSvgChildren(
            Graphics g,
            XmlNode parent,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            bool drew = false;

            for (int i = 0; i < parent.ChildNodes.Count; i++)
            {
                XmlNode node = parent.ChildNodes[i];
                string name = node.LocalName;

                if (name == "defs" ||
                    name == "linearGradient" ||
                    name == "radialGradient" ||
                    name == "clipPath" ||
                    name == "pattern" ||
                    name == "mask" ||
                    name == "filter")
                {
                    continue;
                }

                if (name == "g" ||
                    name == "svg")
                {
                    GraphicsState groupState = g.Save();
                    try
                    {
                        ApplySimpleSvgTransform(
                            g,
                            node,
                            target,
                            minX,
                            minY,
                            sx,
                            sy);

                        drew =
                            DrawEnhancedSvgChildren(
                                g,
                                node,
                                target,
                                minX,
                                minY,
                                sx,
                                sy) ||
                            drew;
                    }
                    finally
                    {
                        g.Restore(groupState);
                    }
                    continue;
                }

                if (name == "text")
                {
                    DrawEnhancedSvgText(
                        g,
                        node,
                        target,
                        minX,
                        minY,
                        sx,
                        sy);
                    drew = true;
                    continue;
                }

                using (GraphicsPath path =
                    BuildEnhancedSvgElementPath(
                        node,
                        target,
                        minX,
                        minY,
                        sx,
                        sy))
                {
                    if (path == null ||
                        path.PointCount == 0)
                    {
                        continue;
                    }

                    ApplySvgFillRule(
                        path,
                        node);

                    GraphicsState state = g.Save();
                    GraphicsPath clip = null;
                    GraphicsPath mask = null;
                    GraphicsPath filteredPath = null;

                    try
                    {
                        ApplySimpleSvgTransform(
                            g,
                            node,
                            target,
                            minX,
                            minY,
                            sx,
                            sy);

                        RectangleF elementBounds =
                            path.GetBounds();

                        clip =
                            BuildSvgClipPath(
                                node,
                                parent.OwnerDocument,
                                elementBounds,
                                target,
                                minX,
                                minY,
                                sx,
                                sy);

                        if (clip != null &&
                            clip.PointCount > 0)
                        {
                            g.SetClip(
                                clip,
                                CombineMode.Intersect);
                        }

                        mask =
                            BuildSvgMaskPath(
                                node,
                                parent.OwnerDocument,
                                elementBounds,
                                target,
                                minX,
                                minY,
                                sx,
                                sy);

                        if (mask != null &&
                            mask.PointCount > 0)
                        {
                            g.SetClip(
                                mask,
                                CombineMode.Intersect);
                        }

                        DrawSvgDropShadowApproximation(
                            g,
                            node,
                            parent.OwnerDocument,
                            path,
                            sx,
                            sy);

                        DrawSvgGaussianBlurApproximation(
                            g,
                            node,
                            parent.OwnerDocument,
                            path,
                            sx,
                            sy);

                        filteredPath =
                            CreateStandaloneSvgOffsetPath(
                                node,
                                parent.OwnerDocument,
                                path,
                                sx,
                                sy);

                        GraphicsPath renderPath =
                            filteredPath ??
                            path;

                        using (Brush fillBrush =
                            CreateSvgFillBrush(
                                node,
                                parent.OwnerDocument,
                                renderPath.GetBounds(),
                                target,
                                minX,
                                minY,
                                sx,
                                sy))
                        {
                            if (fillBrush != null)
                                g.FillPath(
                                    fillBrush,
                                    renderPath);
                        }

                        Color stroke =
                            ReadSvgColorInherited(
                                node,
                                "stroke",
                                Color.Transparent);

                        float strokeOpacity =
                            ReadSvgOpacity(
                                GetSvgStyleInherited(
                                    node,
                                    "stroke-opacity"),
                                1f);

                        stroke =
                            Color.FromArgb(
                                Math.Max(
                                    0,
                                    Math.Min(
                                        255,
                                        (int)Math.Round(
                                            stroke.A *
                                            strokeOpacity))),
                                stroke);

                        float strokeWidth =
                            Math.Max(
                                1f,
                                ParseSvgFloat(
                                    GetSvgStyleInherited(
                                        node,
                                        "stroke-width"),
                                    1f) *
                                Math.Min(
                                    sx,
                                    sy));

                        if (stroke.A > 0)
                        {
                            using (Pen pen =
                                new Pen(
                                    stroke,
                                    strokeWidth))
                            {
                                string lineCap =
                                    GetSvgStyleInherited(
                                        node,
                                        "stroke-linecap");

                                if (lineCap == "round")
                                    pen.StartCap = pen.EndCap = LineCap.Round;
                                else if (lineCap == "square")
                                    pen.StartCap = pen.EndCap = LineCap.Square;

                                string lineJoin =
                                    GetSvgStyleInherited(
                                        node,
                                        "stroke-linejoin");

                                if (lineJoin == "round")
                                    pen.LineJoin = LineJoin.Round;
                                else if (lineJoin == "bevel")
                                    pen.LineJoin = LineJoin.Bevel;

                                ApplySvgStrokeMiterLimit(
                                    pen,
                                    node);

                                ApplySvgStrokeDashPattern(
                                    pen,
                                    node,
                                    target,
                                    sx,
                                    sy,
                                    strokeWidth);

                                g.DrawPath(
                                    pen,
                                    renderPath);
                            }
                        }

                        drew = true;
                    }
                    finally
                    {
                        if (clip != null)
                            clip.Dispose();

                        if (mask != null)
                            mask.Dispose();

                        if (filteredPath != null)
                            filteredPath.Dispose();

                        g.Restore(state);
                    }
                }
            }

            if (!drew)
            {
                return TryDrawSimpleSvgFromDocument(
                    g,
                    parent.OwnerDocument,
                    target);
            }

            return true;
        }

        private static GraphicsPath BuildEnhancedSvgElementPath(
            XmlNode node,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            if (node == null)
                return null;

            string name = node.LocalName;

            if (name == "path")
            {
                return BuildSvgPath(
                    GetAttr(node, "d"),
                    target,
                    minX,
                    minY,
                    sx,
                    sy);
            }

            GraphicsPath path =
                new GraphicsPath();

            if (name == "rect")
            {
                float x =
                    SvgX(
                        target,
                        minX,
                        sx,
                        ParseSvgFloat(
                            GetAttr(node, "x"),
                            0f));

                float y =
                    SvgY(
                        target,
                        minY,
                        sy,
                        ParseSvgFloat(
                            GetAttr(node, "y"),
                            0f));

                float width =
                    Math.Max(
                        0f,
                        ParseSvgFloat(
                            GetAttr(node, "width"),
                            0f) *
                        sx);

                float height =
                    Math.Max(
                        0f,
                        ParseSvgFloat(
                            GetAttr(node, "height"),
                            0f) *
                        sy);

                float rx =
                    Math.Max(
                        0f,
                        ParseSvgFloat(
                            GetAttr(node, "rx"),
                            0f) *
                        sx);

                float ry =
                    Math.Max(
                        0f,
                        ParseSvgFloat(
                            GetAttr(node, "ry"),
                            0f) *
                        sy);

                RectangleF rect =
                    new RectangleF(
                        x,
                        y,
                        width,
                        height);

                if (rx > 0f || ry > 0f)
                {
                    float radius =
                        Math.Min(
                            Math.Max(
                                rx,
                                ry),
                            Math.Min(
                                width,
                                height) /
                            2f);

                    using (GraphicsPath rounded =
                        RoundedRectanglePath(
                            rect,
                            radius))
                    {
                        path.AddPath(
                            rounded,
                            false);
                    }
                }
                else
                {
                    path.AddRectangle(rect);
                }

                return path;
            }

            if (name == "circle" ||
                name == "ellipse")
            {
                float cx =
                    SvgX(
                        target,
                        minX,
                        sx,
                        ParseSvgFloat(
                            GetAttr(node, "cx"),
                            0f));

                float cy =
                    SvgY(
                        target,
                        minY,
                        sy,
                        ParseSvgFloat(
                            GetAttr(node, "cy"),
                            0f));

                float rx =
                    name == "circle"
                        ? ParseSvgFloat(
                            GetAttr(node, "r"),
                            0f) *
                            Math.Min(
                                sx,
                                sy)
                        : ParseSvgFloat(
                            GetAttr(node, "rx"),
                            0f) *
                            sx;

                float ry =
                    name == "circle"
                        ? rx
                        : ParseSvgFloat(
                            GetAttr(node, "ry"),
                            0f) *
                            sy;

                path.AddEllipse(
                    new RectangleF(
                        cx - rx,
                        cy - ry,
                        rx * 2f,
                        ry * 2f));

                return path;
            }

            if (name == "line")
            {
                path.AddLine(
                    SvgX(
                        target,
                        minX,
                        sx,
                        ParseSvgFloat(
                            GetAttr(node, "x1"),
                            0f)),
                    SvgY(
                        target,
                        minY,
                        sy,
                        ParseSvgFloat(
                            GetAttr(node, "y1"),
                            0f)),
                    SvgX(
                        target,
                        minX,
                        sx,
                        ParseSvgFloat(
                            GetAttr(node, "x2"),
                            0f)),
                    SvgY(
                        target,
                        minY,
                        sy,
                        ParseSvgFloat(
                            GetAttr(node, "y2"),
                            0f)));

                return path;
            }

            if (name == "polygon" ||
                name == "polyline")
            {
                List<PointF> points =
                    ParseSvgPoints(
                        GetAttr(
                            node,
                            "points"),
                        target,
                        minX,
                        minY,
                        sx,
                        sy);

                if (points.Count >= 2)
                {
                    if (name == "polygon" &&
                        points.Count >= 3)
                    {
                        path.AddPolygon(
                            points.ToArray());
                    }
                    else
                    {
                        path.AddLines(
                            points.ToArray());
                    }
                }

                return path;
            }

            path.Dispose();
            return null;
        }

        private static void DrawEnhancedSvgText(
            Graphics g,
            XmlNode node,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            float x =
                SvgX(
                    target,
                    minX,
                    sx,
                    ParseSvgFloat(
                        GetAttr(node, "x"),
                        0f));

            float y =
                SvgY(
                    target,
                    minY,
                    sy,
                    ParseSvgFloat(
                        GetAttr(node, "y"),
                        0f));

            float fontSize =
                Math.Max(
                    6f,
                    ParseSvgFloat(
                        GetSvgStyleInherited(
                            node,
                            "font-size"),
                        12f) *
                    sy);

            string family =
                GetSvgStyleInherited(
                    node,
                    "font-family");

            if (string.IsNullOrEmpty(family))
                family = "Arial";

            family =
                family
                    .Trim()
                    .Trim(
                        '"',
                        '\'');

            Color fill =
                ReadSvgColorInherited(
                    node,
                    "fill",
                    Color.Black);

            float opacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "opacity"),
                    1f) *
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "fill-opacity"),
                    1f);

            fill =
                Color.FromArgb(
                    Math.Max(
                        0,
                        Math.Min(
                            255,
                            (int)Math.Round(
                                fill.A *
                                opacity))),
                    fill);

            using (Font font =
                SafeFont(
                    family,
                    fontSize))
            using (Brush brush =
                new SolidBrush(fill))
            {
                g.DrawString(
                    node.InnerText,
                    font,
                    brush,
                    x,
                    y - font.Height);
            }
        }

        private static Brush CreateSvgFillBrush(
            XmlNode node,
            XmlDocument document,
            RectangleF bounds,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            string raw =
                GetSvgStyleInherited(
                    node,
                    "fill");

            if (string.IsNullOrEmpty(raw))
                raw = "black";

            raw = raw.Trim();

            if (string.Equals(
                    raw,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            float opacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "opacity"),
                    1f) *
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "fill-opacity"),
                    1f);

            string gradientId =
                ExtractSvgUrlId(raw);

            if (!string.IsNullOrEmpty(
                    gradientId))
            {
                XmlNode gradient =
                    FindSvgNodeById(
                        document,
                        gradientId);

                if (gradient != null)
                {
                    if (gradient.LocalName ==
                        "pattern")
                    {
                        Brush patternBrush =
                            CreateSvgPatternBrush(
                                gradient,
                                document,
                                bounds,
                                target,
                                minX,
                                minY,
                                sx,
                                sy,
                                opacity);

                        if (patternBrush != null)
                            return patternBrush;
                    }
                    else
                    {
                        Brush gradientBrush =
                            CreateSvgGradientBrush(
                                gradient,
                                document,
                                bounds,
                                target,
                                minX,
                                minY,
                                sx,
                                sy,
                                opacity);

                        if (gradientBrush != null)
                            return gradientBrush;
                    }
                }
            }

            Color color =
                ReadSvgColorInherited(
                    node,
                    "fill",
                    Color.Black);

            color =
                Color.FromArgb(
                    Math.Max(
                        0,
                        Math.Min(
                            255,
                            (int)Math.Round(
                                color.A *
                                opacity))),
                    color);

            return color.A > 0
                ? (Brush)new SolidBrush(
                    color)
                : null;
        }

        private static Brush CreateSvgPatternBrush(
            XmlNode pattern,
            XmlDocument document,
            RectangleF bounds,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy,
            float opacity)
        {
            if (pattern == null)
                return null;

            bool userSpace =
                string.Equals(
                    GetAttr(
                        pattern,
                        "patternUnits"),
                    "userSpaceOnUse",
                    StringComparison.OrdinalIgnoreCase);

            float logicalWidth =
                ReadSvgPatternLogicalLength(
                    GetAttr(
                        pattern,
                        "width"),
                    userSpace
                        ? 16f
                        : 0.125f);

            float logicalHeight =
                ReadSvgPatternLogicalLength(
                    GetAttr(
                        pattern,
                        "height"),
                    userSpace
                        ? 16f
                        : 0.125f);

            float tileWidth =
                ResolveSvgPatternPixelLength(
                    GetAttr(
                        pattern,
                        "width"),
                    bounds.Width,
                    sx,
                    userSpace,
                    Math.Max(
                        2f,
                        bounds.Width /
                        8f));

            float tileHeight =
                ResolveSvgPatternPixelLength(
                    GetAttr(
                        pattern,
                        "height"),
                    bounds.Height,
                    sy,
                    userSpace,
                    Math.Max(
                        2f,
                        bounds.Height /
                        8f));

            int bitmapWidth =
                Math.Max(
                    2,
                    Math.Min(
                        512,
                        (int)Math.Ceiling(
                            tileWidth)));

            int bitmapHeight =
                Math.Max(
                    2,
                    Math.Min(
                        512,
                        (int)Math.Ceiling(
                            tileHeight)));

            using (Bitmap tile =
                new Bitmap(
                    bitmapWidth,
                    bitmapHeight,
                    PixelFormat.Format32bppArgb))
            {
                using (Graphics tileGraphics =
                    Graphics.FromImage(tile))
                {
                    tileGraphics.Clear(
                        Color.Transparent);
                    tileGraphics.SmoothingMode =
                        SmoothingMode.AntiAlias;
                    tileGraphics.InterpolationMode =
                        InterpolationMode.HighQualityBicubic;

                    RectangleF tileTarget =
                        new RectangleF(
                            0f,
                            0f,
                            bitmapWidth,
                            bitmapHeight);

                    float patternMinX = 0f;
                    float patternMinY = 0f;
                    float patternSx;
                    float patternSy;

                    List<float> viewBox =
                        ParseSvgNumberList(
                            GetAttr(
                                pattern,
                                "viewBox"));

                    if (viewBox.Count >= 4 &&
                        Math.Abs(
                            viewBox[2]) >
                        0.0001f &&
                        Math.Abs(
                            viewBox[3]) >
                        0.0001f)
                    {
                        patternMinX =
                            viewBox[0];
                        patternMinY =
                            viewBox[1];
                        patternSx =
                            bitmapWidth /
                            viewBox[2];
                        patternSy =
                            bitmapHeight /
                            viewBox[3];
                    }
                    else
                    {
                        bool contentObjectBox =
                            string.Equals(
                                GetAttr(
                                    pattern,
                                    "patternContentUnits"),
                                "objectBoundingBox",
                                StringComparison.OrdinalIgnoreCase);

                        if (contentObjectBox)
                        {
                            patternSx =
                                bitmapWidth;
                            patternSy =
                                bitmapHeight;
                        }
                        else
                        {
                            patternSx =
                                bitmapWidth /
                                Math.Max(
                                    0.0001f,
                                    logicalWidth);
                            patternSy =
                                bitmapHeight /
                                Math.Max(
                                    0.0001f,
                                    logicalHeight);
                        }
                    }

                    DrawEnhancedSvgChildren(
                        tileGraphics,
                        pattern,
                        tileTarget,
                        patternMinX,
                        patternMinY,
                        patternSx,
                        patternSy);
                }

                if (opacity < 0.999f)
                {
                    MultiplyBitmapAlpha(
                        tile,
                        opacity);
                }

                TextureBrush brush =
                    new TextureBrush(
                        tile,
                        WrapMode.Tile);

                float offsetX =
                    ResolveSvgPatternOffset(
                        GetAttr(
                            pattern,
                            "x"),
                        bounds.Left,
                        bounds.Width,
                        target.Left,
                        minX,
                        sx,
                        userSpace);

                float offsetY =
                    ResolveSvgPatternOffset(
                        GetAttr(
                            pattern,
                            "y"),
                        bounds.Top,
                        bounds.Height,
                        target.Top,
                        minY,
                        sy,
                        userSpace);

                brush.TranslateTransform(
                    offsetX,
                    offsetY,
                    MatrixOrder.Append);

                string transform =
                    GetAttr(
                        pattern,
                        "patternTransform");

                if (!string.IsNullOrEmpty(
                        transform))
                {
                    using (Matrix matrix =
                        BuildSvgBrushTransformMatrix(
                            transform,
                            bounds,
                            target,
                            minX,
                            minY,
                            sx,
                            sy,
                            userSpace))
                    {
                        if (matrix != null)
                        {
                            brush.MultiplyTransform(
                                matrix,
                                MatrixOrder.Append);
                        }
                    }
                }

                return brush;
            }
        }

        private static float ReadSvgPatternLogicalLength(
            string raw,
            float fallback)
        {
            if (string.IsNullOrEmpty(raw))
                return fallback;

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                return Math.Max(
                    0.0001f,
                    ParseSvgFloat(
                        raw.Substring(
                            0,
                            raw.Length - 1),
                        fallback *
                        100f) /
                    100f);
            }

            return Math.Max(
                0.0001f,
                ParseSvgFloat(
                    raw,
                    fallback));
        }

        private static float ResolveSvgPatternPixelLength(
            string raw,
            float reference,
            float scale,
            bool userSpace,
            float fallback)
        {
            if (string.IsNullOrEmpty(raw))
                return fallback;

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                return Math.Max(
                    2f,
                    reference *
                    ParseSvgFloat(
                        raw.Substring(
                            0,
                            raw.Length - 1),
                        12.5f) /
                    100f);
            }

            float value =
                ParseSvgFloat(
                    raw,
                    userSpace
                        ? 16f
                        : 0.125f);

            return Math.Max(
                2f,
                userSpace
                    ? Math.Abs(
                        value *
                        scale)
                    : Math.Abs(
                        value *
                        reference));
        }

        private static float ResolveSvgPatternOffset(
            string raw,
            float boundsStart,
            float boundsLength,
            float targetStart,
            float viewMin,
            float scale,
            bool userSpace)
        {
            if (string.IsNullOrEmpty(raw))
                return 0f;

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                return boundsStart +
                    boundsLength *
                    ParseSvgFloat(
                        raw.Substring(
                            0,
                            raw.Length - 1),
                        0f) /
                    100f;
            }

            float value =
                ParseSvgFloat(
                    raw,
                    0f);

            if (userSpace)
            {
                return targetStart +
                    (value -
                     viewMin) *
                    scale;
            }

            return boundsStart +
                value *
                boundsLength;
        }

        private static void MultiplyBitmapAlpha(
            Bitmap bitmap,
            float opacity)
        {
            if (bitmap == null)
                return;

            opacity =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        opacity));

            for (int y = 0;
                 y < bitmap.Height;
                 y++)
            {
                for (int x = 0;
                     x < bitmap.Width;
                     x++)
                {
                    Color color =
                        bitmap.GetPixel(
                            x,
                            y);

                    bitmap.SetPixel(
                        x,
                        y,
                        Color.FromArgb(
                            Math.Max(
                                0,
                                Math.Min(
                                    255,
                                    (int)Math.Round(
                                        color.A *
                                        opacity))),
                            color.R,
                            color.G,
                            color.B));
                }
            }
        }

        private static Brush CreateSvgGradientBrush(
            XmlNode gradient,
            XmlDocument document,
            RectangleF bounds,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy,
            float opacity)
        {
            List<SvgGradientStop> stops =
                ReadSvgGradientStops(
                    gradient,
                    document,
                    opacity);

            if (stops.Count == 0)
                return null;

            if (stops.Count == 1)
            {
                return new SolidBrush(
                    stops[0].Color);
            }

            ColorBlend blend =
                BuildSvgColorBlend(
                    stops);

            bool userSpace =
                string.Equals(
                    GetAttr(
                        gradient,
                        "gradientUnits"),
                    "userSpaceOnUse",
                    StringComparison.OrdinalIgnoreCase);

            if (gradient.LocalName ==
                "radialGradient")
            {
                float cx =
                    ResolveSvgGradientCoordinate(
                        GetAttr(
                            gradient,
                            "cx"),
                        bounds.Left,
                        bounds.Width,
                        target.Left,
                        minX,
                        sx,
                        0.5f,
                        userSpace);

                float cy =
                    ResolveSvgGradientCoordinate(
                        GetAttr(
                            gradient,
                            "cy"),
                        bounds.Top,
                        bounds.Height,
                        target.Top,
                        minY,
                        sy,
                        0.5f,
                        userSpace);

                float radius =
                    ResolveSvgGradientRadius(
                        GetAttr(
                            gradient,
                            "r"),
                        bounds,
                        Math.Min(
                            sx,
                            sy),
                        0.5f,
                        userSpace);

                radius =
                    Math.Max(
                        1f,
                        radius);

                using (GraphicsPath ellipse =
                    new GraphicsPath())
                {
                    ellipse.AddEllipse(
                        new RectangleF(
                            cx - radius,
                            cy - radius,
                            radius * 2f,
                            radius * 2f));

                    PathGradientBrush brush =
                        new PathGradientBrush(
                            ellipse);

                    brush.CenterPoint =
                        new PointF(
                            cx,
                            cy);
                    brush.InterpolationColors =
                        blend;

                    ApplySvgGradientSpreadMethod(
                        brush,
                        gradient);

                    ApplySvgGradientTransform(
                        brush,
                        gradient,
                        bounds,
                        target,
                        minX,
                        minY,
                        sx,
                        sy,
                        userSpace);

                    return brush;
                }
            }

            float x1 =
                ResolveSvgGradientCoordinate(
                    GetAttr(
                        gradient,
                        "x1"),
                    bounds.Left,
                    bounds.Width,
                    target.Left,
                    minX,
                    sx,
                    0f,
                    userSpace);

            float y1 =
                ResolveSvgGradientCoordinate(
                    GetAttr(
                        gradient,
                        "y1"),
                    bounds.Top,
                    bounds.Height,
                    target.Top,
                    minY,
                    sy,
                    0f,
                    userSpace);

            float x2 =
                ResolveSvgGradientCoordinate(
                    GetAttr(
                        gradient,
                        "x2"),
                    bounds.Left,
                    bounds.Width,
                    target.Left,
                    minX,
                    sx,
                    1f,
                    userSpace);

            float y2 =
                ResolveSvgGradientCoordinate(
                    GetAttr(
                        gradient,
                        "y2"),
                    bounds.Top,
                    bounds.Height,
                    target.Top,
                    minY,
                    sy,
                    0f,
                    userSpace);

            if (Math.Abs(x2 - x1) < 0.01f &&
                Math.Abs(y2 - y1) < 0.01f)
            {
                x2 = x1 + 1f;
            }

            LinearGradientBrush linear =
                new LinearGradientBrush(
                    new PointF(
                        x1,
                        y1),
                    new PointF(
                        x2,
                        y2),
                    stops[0].Color,
                    stops[
                        stops.Count - 1]
                        .Color);

            linear.InterpolationColors =
                blend;

            ApplySvgGradientSpreadMethod(
                linear,
                gradient);

            ApplySvgGradientTransform(
                linear,
                gradient,
                bounds,
                target,
                minX,
                minY,
                sx,
                sy,
                userSpace);

            return linear;
        }

        private static List<SvgGradientStop>
            ReadSvgGradientStops(
                XmlNode gradient,
                XmlDocument document,
                float opacity)
        {
            List<SvgGradientStop> result =
                new List<SvgGradientStop>();

            if (gradient == null)
                return result;

            for (int i = 0;
                 i < gradient.ChildNodes.Count;
                 i++)
            {
                XmlNode stop =
                    gradient.ChildNodes[i];

                if (stop.LocalName != "stop")
                    continue;

                float offset =
                    ReadSvgGradientOffset(
                        GetAttr(
                            stop,
                            "offset"));

                Color color =
                    ReadSvgColor(
                        stop,
                        "stop-color",
                        Color.Black);

                float stopOpacity =
                    ReadSvgOpacity(
                        GetSvgStyle(
                            stop,
                            "stop-opacity"),
                        1f) *
                    opacity;

                color =
                    Color.FromArgb(
                        Math.Max(
                            0,
                            Math.Min(
                                255,
                                (int)Math.Round(
                                    color.A *
                                    stopOpacity))),
                        color);

                SvgGradientStop item =
                    new SvgGradientStop();
                item.Offset = offset;
                item.Color = color;
                result.Add(item);
            }

            if (result.Count == 0)
            {
                string href = null;

                if (gradient.Attributes != null)
                {
                    for (int i = 0;
                         i < gradient.Attributes.Count;
                         i++)
                    {
                        XmlAttribute attribute =
                            gradient.Attributes[i];

                        if (attribute.LocalName ==
                            "href")
                        {
                            href =
                                attribute.Value;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(href) &&
                    href.StartsWith("#"))
                {
                    XmlNode inherited =
                        FindSvgNodeById(
                            document,
                            href.Substring(1));

                    if (inherited != null &&
                        !object.ReferenceEquals(
                            inherited,
                            gradient))
                    {
                        return ReadSvgGradientStops(
                            inherited,
                            document,
                            opacity);
                    }
                }
            }

            result.Sort(
                delegate(
                    SvgGradientStop left,
                    SvgGradientStop right)
                {
                    return left.Offset
                        .CompareTo(
                            right.Offset);
                });

            if (result.Count > 0 &&
                result[0].Offset > 0f)
            {
                SvgGradientStop first =
                    new SvgGradientStop();
                first.Offset = 0f;
                first.Color =
                    result[0].Color;
                result.Insert(
                    0,
                    first);
            }

            if (result.Count > 0 &&
                result[
                    result.Count - 1]
                    .Offset < 1f)
            {
                SvgGradientStop last =
                    new SvgGradientStop();
                last.Offset = 1f;
                last.Color =
                    result[
                        result.Count - 1]
                        .Color;
                result.Add(last);
            }

            return result;
        }

        private static ColorBlend BuildSvgColorBlend(
            List<SvgGradientStop> stops)
        {
            ColorBlend blend =
                new ColorBlend(
                    stops.Count);

            Color[] colors =
                new Color[
                    stops.Count];
            float[] positions =
                new float[
                    stops.Count];

            float previous = 0f;

            for (int i = 0;
                 i < stops.Count;
                 i++)
            {
                float offset =
                    Math.Max(
                        previous,
                        Math.Min(
                            1f,
                            stops[i].Offset));

                colors[i] =
                    stops[i].Color;
                positions[i] =
                    offset;
                previous = offset;
            }

            positions[0] = 0f;
            positions[
                positions.Length - 1] =
                1f;

            blend.Colors =
                colors;
            blend.Positions =
                positions;
            return blend;
        }

        private static GraphicsPath BuildSvgClipPath(
            XmlNode node,
            XmlDocument document,
            RectangleF bounds,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            string raw =
                GetSvgStyleInherited(
                    node,
                    "clip-path");

            string id =
                ExtractSvgUrlId(raw);

            if (string.IsNullOrEmpty(id))
                return null;

            XmlNode clipNode =
                FindSvgNodeById(
                    document,
                    id);

            if (clipNode == null ||
                clipNode.LocalName != "clipPath")
            {
                return null;
            }

            bool objectBox =
                string.Equals(
                    GetAttr(
                        clipNode,
                        "clipPathUnits"),
                    "objectBoundingBox",
                    StringComparison.OrdinalIgnoreCase);

            RectangleF pathTarget =
                objectBox
                    ? bounds
                    : target;

            float pathMinX =
                objectBox
                    ? 0f
                    : minX;
            float pathMinY =
                objectBox
                    ? 0f
                    : minY;
            float pathSx =
                objectBox
                    ? bounds.Width
                    : sx;
            float pathSy =
                objectBox
                    ? bounds.Height
                    : sy;

            GraphicsPath result =
                new GraphicsPath();

            result.FillMode =
                FillMode.Winding;

            AppendSvgGeometryChildren(
                result,
                clipNode,
                pathTarget,
                pathMinX,
                pathMinY,
                pathSx,
                pathSy,
                false,
                null);

            if (result.PointCount == 0)
            {
                result.Dispose();
                return null;
            }

            return result;
        }

        private static GraphicsPath BuildSvgMaskPath(
            XmlNode node,
            XmlDocument document,
            RectangleF bounds,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            string raw =
                GetSvgStyleInherited(
                    node,
                    "mask");

            string id =
                ExtractSvgUrlId(raw);

            if (string.IsNullOrEmpty(id))
                return null;

            XmlNode maskNode =
                FindSvgNodeById(
                    document,
                    id);

            if (maskNode == null ||
                maskNode.LocalName != "mask")
            {
                return null;
            }

            bool objectBox =
                string.Equals(
                    GetAttr(
                        maskNode,
                        "maskContentUnits"),
                    "objectBoundingBox",
                    StringComparison.OrdinalIgnoreCase);

            RectangleF pathTarget =
                objectBox
                    ? bounds
                    : target;

            float pathMinX =
                objectBox
                    ? 0f
                    : minX;
            float pathMinY =
                objectBox
                    ? 0f
                    : minY;
            float pathSx =
                objectBox
                    ? bounds.Width
                    : sx;
            float pathSy =
                objectBox
                    ? bounds.Height
                    : sy;

            GraphicsPath result =
                new GraphicsPath();

            AppendSvgGeometryChildren(
                result,
                maskNode,
                pathTarget,
                pathMinX,
                pathMinY,
                pathSx,
                pathSy,
                true,
                null);

            if (result.PointCount == 0)
            {
                result.Dispose();
                return null;
            }

            return result;
        }

        private static void AppendSvgGeometryChildren(
            GraphicsPath destination,
            XmlNode parent,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy,
            bool maskMode,
            Matrix inheritedTransform)
        {
            if (destination == null ||
                parent == null)
            {
                return;
            }

            for (int i = 0;
                 i < parent.ChildNodes.Count;
                 i++)
            {
                XmlNode child =
                    parent.ChildNodes[i];

                Matrix combined =
                    inheritedTransform != null
                        ? inheritedTransform.Clone()
                        : new Matrix();

                try
                {
                    string transform =
                        GetAttr(
                            child,
                            "transform");

                    if (!string.IsNullOrEmpty(
                            transform))
                    {
                        using (Matrix local =
                            BuildSvgGeometryTransformMatrix(
                                transform,
                                target,
                                minX,
                                minY,
                                sx,
                                sy))
                        {
                            if (local != null)
                            {
                                combined.Multiply(
                                    local,
                                    MatrixOrder.Append);
                            }
                        }
                    }

                    if (child.LocalName == "g")
                    {
                        AppendSvgGeometryChildren(
                            destination,
                            child,
                            target,
                            minX,
                            minY,
                            sx,
                            sy,
                            maskMode,
                            combined);
                        continue;
                    }

                    if (maskMode &&
                        !SvgMaskNodeIsVisible(
                            child))
                    {
                        continue;
                    }

                    using (GraphicsPath childPath =
                        BuildEnhancedSvgElementPath(
                            child,
                            target,
                            minX,
                            minY,
                            sx,
                            sy))
                    {
                        if (childPath != null &&
                            childPath.PointCount > 0)
                        {
                            if (!combined.IsIdentity)
                            {
                                childPath.Transform(
                                    combined);
                            }

                            if (!maskMode)
                            {
                                string clipRule =
                                    GetSvgStyleInherited(
                                        child,
                                        "clip-rule");

                                if (string.Equals(
                                        clipRule,
                                        "evenodd",
                                        StringComparison.OrdinalIgnoreCase))
                                {
                                    destination.FillMode =
                                        FillMode.Alternate;
                                }
                                else if (string.Equals(
                                        clipRule,
                                        "nonzero",
                                        StringComparison.OrdinalIgnoreCase))
                                {
                                    destination.FillMode =
                                        FillMode.Winding;
                                }
                            }

                            destination.AddPath(
                                childPath,
                                false);
                        }
                    }
                }
                finally
                {
                    combined.Dispose();
                }
            }
        }

        private static Matrix BuildSvgGeometryTransformMatrix(
            string transform,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            Matrix result =
                new Matrix();

            if (string.IsNullOrEmpty(transform))
                return result;

            int index = 0;

            while (index < transform.Length)
            {
                while (index < transform.Length &&
                    (char.IsWhiteSpace(
                        transform[index]) ||
                     transform[index] == ','))
                {
                    index++;
                }

                int nameStart =
                    index;

                while (index < transform.Length &&
                    char.IsLetter(
                        transform[index]))
                {
                    index++;
                }

                if (index <= nameStart)
                {
                    index++;
                    continue;
                }

                string name =
                    transform.Substring(
                        nameStart,
                        index -
                        nameStart)
                    .ToLowerInvariant();

                while (index < transform.Length &&
                    char.IsWhiteSpace(
                        transform[index]))
                {
                    index++;
                }

                if (index >= transform.Length ||
                    transform[index] != '(')
                {
                    continue;
                }

                int argumentStart =
                    index + 1;
                int close =
                    transform.IndexOf(
                        ')',
                        argumentStart);

                if (close < 0)
                    break;

                List<float> values =
                    ParseSvgNumberList(
                        transform.Substring(
                            argumentStart,
                            close -
                            argumentStart));

                if (name == "matrix" &&
                    values.Count >= 6)
                {
                    float safeSx =
                        Math.Abs(sx) <
                        0.0001f
                            ? 1f
                            : sx;
                    float safeSy =
                        Math.Abs(sy) <
                        0.0001f
                            ? 1f
                            : sy;

                    using (Matrix operation =
                        new Matrix(
                            values[0],
                            values[1] *
                                safeSy /
                                safeSx,
                            values[2] *
                                safeSx /
                                safeSy,
                            values[3],
                            values[4] *
                                sx,
                            values[5] *
                                sy))
                    {
                        result.Multiply(
                            operation,
                            MatrixOrder.Append);
                    }
                }
                else if (name == "translate" &&
                    values.Count > 0)
                {
                    result.Translate(
                        values[0] *
                            sx,
                        (values.Count > 1
                            ? values[1]
                            : 0f) *
                            sy,
                        MatrixOrder.Append);
                }
                else if (name == "scale" &&
                    values.Count > 0)
                {
                    result.Scale(
                        values[0],
                        values.Count > 1
                            ? values[1]
                            : values[0],
                        MatrixOrder.Append);
                }
                else if (name == "rotate" &&
                    values.Count > 0)
                {
                    if (values.Count >= 3)
                    {
                        float cx =
                            SvgX(
                                target,
                                minX,
                                sx,
                                values[1]);
                        float cy =
                            SvgY(
                                target,
                                minY,
                                sy,
                                values[2]);

                        result.Translate(
                            cx,
                            cy,
                            MatrixOrder.Append);
                        result.Rotate(
                            values[0],
                            MatrixOrder.Append);
                        result.Translate(
                            -cx,
                            -cy,
                            MatrixOrder.Append);
                    }
                    else
                    {
                        result.Rotate(
                            values[0],
                            MatrixOrder.Append);
                    }
                }
                else if (name == "skewx" &&
                    values.Count > 0)
                {
                    float tangent =
                        (float)Math.Tan(
                            values[0] *
                            Math.PI /
                            180.0);

                    using (Matrix operation =
                        new Matrix(
                            1f,
                            0f,
                            tangent,
                            1f,
                            0f,
                            0f))
                    {
                        result.Multiply(
                            operation,
                            MatrixOrder.Append);
                    }
                }
                else if (name == "skewy" &&
                    values.Count > 0)
                {
                    float tangent =
                        (float)Math.Tan(
                            values[0] *
                            Math.PI /
                            180.0);

                    using (Matrix operation =
                        new Matrix(
                            1f,
                            tangent,
                            0f,
                            1f,
                            0f,
                            0f))
                    {
                        result.Multiply(
                            operation,
                            MatrixOrder.Append);
                    }
                }

                index =
                    close + 1;
            }

            return result;
        }

        private static bool SvgMaskNodeIsVisible(
            XmlNode node)
        {
            if (node == null)
                return false;

            string fill =
                GetSvgStyleInherited(
                    node,
                    "fill");

            if (string.Equals(
                    fill,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            float opacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "opacity"),
                    1f) *
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "fill-opacity"),
                    1f);

            if (opacity <= 0.01f)
                return false;

            if (!string.IsNullOrEmpty(fill))
            {
                Color color =
                    ReadSvgColorInherited(
                        node,
                        "fill",
                        Color.White);

                int luminance =
                    color.R +
                    color.G +
                    color.B;

                if (color.A <= 5 ||
                    luminance <= 12)
                {
                    return false;
                }
            }

            return true;
        }

        private static void ApplySvgFillRule(
            GraphicsPath path,
            XmlNode node)
        {
            if (path == null)
            {
                return;
            }

            string rule =
                node == null
                    ? null
                    : GetSvgStyleInherited(
                        node,
                        "fill-rule");

            path.FillMode =
                string.Equals(
                    rule,
                    "evenodd",
                    StringComparison.OrdinalIgnoreCase)
                    ? FillMode.Alternate
                    : FillMode.Winding;
        }

        private static void ApplySvgStrokeMiterLimit(
            Pen pen,
            XmlNode node)
        {
            if (pen == null)
            {
                return;
            }

            float limit =
                4f;

            if (node != null)
            {
                string raw =
                    GetSvgStyleInherited(
                        node,
                        "stroke-miterlimit");

                if (!string.IsNullOrEmpty(
                        raw))
                {
                    float parsed =
                        ParseSvgFloat(
                            raw,
                            4f);

                    if (!float.IsNaN(parsed) &&
                        !float.IsInfinity(parsed) &&
                        parsed >= 1f)
                    {
                        limit =
                            parsed;
                    }
                }
            }

            pen.MiterLimit =
                limit;
        }

        private static bool ApplySvgStrokeDashPattern(
            Pen pen,
            XmlNode node,
            RectangleF target,
            float sx,
            float sy,
            float strokeWidth)
        {
            if (pen == null ||
                node == null ||
                strokeWidth <= 0f)
            {
                return false;
            }

            string raw =
                GetSvgStyleInherited(
                    node,
                    "stroke-dasharray");

            if (string.IsNullOrEmpty(raw) ||
                string.Equals(
                    raw.Trim(),
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string[] tokens =
                raw.Replace(
                    ',',
                    ' ')
                .Split(
                    new char[]
                    {
                        ' ',
                        '\t',
                        '\r',
                        '\n'
                    },
                    StringSplitOptions.RemoveEmptyEntries);

            if (tokens.Length == 0)
            {
                return false;
            }

            float scale =
                Math.Max(
                    0.0001f,
                    Math.Min(
                        Math.Abs(sx),
                        Math.Abs(sy)));

            float normalizedDiagonal =
                (float)(
                    Math.Sqrt(
                        target.Width *
                        target.Width +
                        target.Height *
                        target.Height) /
                    Math.Sqrt(2.0));

            List<float> values =
                new List<float>();

            for (int i = 0;
                 i < tokens.Length;
                 i++)
            {
                string token =
                    tokens[i].Trim();

                if (token.Length == 0)
                    continue;

                float pixels;

                if (token.EndsWith(
                        "%",
                        StringComparison.Ordinal))
                {
                    float percent =
                        ParseSvgFloat(
                            token.Substring(
                                0,
                                token.Length - 1),
                            -1f);

                    if (percent < 0f)
                        return false;

                    pixels =
                        normalizedDiagonal *
                        percent /
                        100f;
                }
                else
                {
                    if (token.EndsWith(
                            "px",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        token =
                            token.Substring(
                                0,
                                token.Length - 2);
                    }

                    float logical =
                        ParseSvgFloat(
                            token,
                            -1f);

                    if (logical < 0f)
                        return false;

                    pixels =
                        logical *
                        scale;
                }

                values.Add(
                    Math.Max(
                        0.05f,
                        pixels /
                        strokeWidth));
            }

            if (values.Count == 0)
            {
                return false;
            }

            int sourceCount =
                values.Count;

            if ((sourceCount % 2) != 0)
            {
                for (int i = 0;
                     i < sourceCount;
                     i++)
                {
                    values.Add(
                        values[i]);
                }
            }

            try
            {
                pen.DashStyle =
                    DashStyle.Custom;
                pen.DashPattern =
                    values.ToArray();

                string rawOffset =
                    GetSvgStyleInherited(
                        node,
                        "stroke-dashoffset");

                if (!string.IsNullOrEmpty(
                        rawOffset))
                {
                    rawOffset =
                        rawOffset.Trim();

                    float offsetPixels;

                    if (rawOffset.EndsWith(
                            "%",
                            StringComparison.Ordinal))
                    {
                        float percent =
                            ParseSvgFloat(
                                rawOffset.Substring(
                                    0,
                                    rawOffset.Length - 1),
                                0f);

                        offsetPixels =
                            normalizedDiagonal *
                            percent /
                            100f;
                    }
                    else
                    {
                        if (rawOffset.EndsWith(
                                "px",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            rawOffset =
                                rawOffset.Substring(
                                    0,
                                    rawOffset.Length - 2);
                        }

                        offsetPixels =
                            ParseSvgFloat(
                                rawOffset,
                                0f) *
                            scale;
                    }

                    pen.DashOffset =
                        offsetPixels /
                        strokeWidth;
                }

                return true;
            }
            catch
            {
                pen.DashStyle =
                    DashStyle.Solid;
                return false;
            }
        }

        private static bool TryReadStandaloneSvgOffset(
            XmlNode node,
            XmlDocument document,
            float sx,
            float sy,
            out float offsetX,
            out float offsetY)
        {
            offsetX = 0f;
            offsetY = 0f;

            if (node == null ||
                document == null)
            {
                return false;
            }

            string filterId =
                ExtractSvgUrlId(
                    GetSvgStyleInherited(
                        node,
                        "filter"));

            if (string.IsNullOrEmpty(
                    filterId))
            {
                return false;
            }

            XmlNode filter =
                FindSvgNodeById(
                    document,
                    filterId);

            if (filter == null ||
                filter.LocalName !=
                    "filter")
            {
                return false;
            }

            XmlNode offset =
                null;
            int primitiveCount =
                0;

            for (int i = 0;
                 i < filter.ChildNodes.Count;
                 i++)
            {
                XmlNode child =
                    filter.ChildNodes[i];

                if (child == null ||
                    child.NodeType !=
                        XmlNodeType.Element)
                {
                    continue;
                }

                string name =
                    child.LocalName;

                if (name == "animate" ||
                    name == "set")
                {
                    continue;
                }

                primitiveCount++;

                if (name == "feOffset")
                {
                    offset =
                        child;
                }
            }

            if (primitiveCount != 1 ||
                offset == null)
            {
                return false;
            }

            string input =
                GetAttr(
                    offset,
                    "in");

            if (!string.IsNullOrEmpty(
                    input) &&
                !string.Equals(
                    input,
                    "SourceGraphic",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            offsetX =
                ParseSvgFloat(
                    GetAttr(
                        offset,
                        "dx"),
                    0f) *
                sx;

            offsetY =
                ParseSvgFloat(
                    GetAttr(
                        offset,
                        "dy"),
                    0f) *
                sy;

            return Math.Abs(offsetX) >
                    0.001f ||
                Math.Abs(offsetY) >
                    0.001f;
        }

        private static GraphicsPath CreateStandaloneSvgOffsetPath(
            XmlNode node,
            XmlDocument document,
            GraphicsPath path,
            float sx,
            float sy)
        {
            if (path == null ||
                path.PointCount == 0)
            {
                return null;
            }

            float offsetX;
            float offsetY;

            if (!TryReadStandaloneSvgOffset(
                    node,
                    document,
                    sx,
                    sy,
                    out offsetX,
                    out offsetY))
            {
                return null;
            }

            GraphicsPath shifted =
                (GraphicsPath)path.Clone();

            using (Matrix translation =
                new Matrix())
            {
                translation.Translate(
                    offsetX,
                    offsetY);

                shifted.Transform(
                    translation);
            }

            return shifted;
        }

        private static bool TryReadSvgDropShadow(
            XmlNode node,
            XmlDocument document,
            float sx,
            float sy,
            out float offsetX,
            out float offsetY,
            out float blurX,
            out float blurY,
            out Color shadowColor)
        {
            offsetX = 0f;
            offsetY = 0f;
            blurX = 0f;
            blurY = 0f;
            shadowColor =
                Color.Black;

            if (node == null ||
                document == null)
            {
                return false;
            }

            string filterValue =
                GetSvgStyleInherited(
                    node,
                    "filter");

            string filterId =
                ExtractSvgUrlId(
                    filterValue);

            if (string.IsNullOrEmpty(
                    filterId))
            {
                return false;
            }

            XmlNode filter =
                FindSvgNodeById(
                    document,
                    filterId);

            if (filter == null ||
                filter.LocalName !=
                    "filter")
            {
                return false;
            }

            XmlNode dropShadow =
                null;

            for (int i = 0;
                 i < filter.ChildNodes.Count;
                 i++)
            {
                XmlNode child =
                    filter.ChildNodes[i];

                if (child != null &&
                    child.LocalName ==
                        "feDropShadow")
                {
                    dropShadow =
                        child;
                    break;
                }
            }

            if (dropShadow == null)
            {
                return false;
            }

            string input =
                GetAttr(
                    dropShadow,
                    "in");

            if (!string.IsNullOrEmpty(
                    input) &&
                !string.Equals(
                    input,
                    "SourceGraphic",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            offsetX =
                ParseSvgFloat(
                    GetAttr(
                        dropShadow,
                        "dx"),
                    2f) *
                sx;

            offsetY =
                ParseSvgFloat(
                    GetAttr(
                        dropShadow,
                        "dy"),
                    2f) *
                sy;

            string rawDeviation =
                GetAttr(
                    dropShadow,
                    "stdDeviation");

            float stdX = 2f;
            float stdY = 2f;

            if (!string.IsNullOrEmpty(
                    rawDeviation))
            {
                string[] parts =
                    rawDeviation.Replace(
                        ',',
                        ' ')
                    .Split(
                        new char[]
                        {
                            ' ',
                            '\t',
                            '\r',
                            '\n'
                        },
                        StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length > 0)
                {
                    stdX =
                        Math.Max(
                            0f,
                            ParseSvgFloat(
                                parts[0],
                                2f));

                    stdY =
                        parts.Length > 1
                            ? Math.Max(
                                0f,
                                ParseSvgFloat(
                                    parts[1],
                                    stdX))
                            : stdX;
                }
            }

            blurX =
                Math.Min(
                    48f,
                    stdX *
                    Math.Abs(sx));

            blurY =
                Math.Min(
                    48f,
                    stdY *
                    Math.Abs(sy));

            Color floodColor =
                ReadSvgColorInherited(
                    dropShadow,
                    "flood-color",
                    Color.Black);

            float floodOpacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        dropShadow,
                        "flood-opacity"),
                    1f);

            int alpha =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            floodColor.A *
                            floodOpacity)));

            shadowColor =
                Color.FromArgb(
                    alpha,
                    floodColor);

            return shadowColor.A > 0;
        }

        private static void DrawSvgDropShadowApproximation(
            Graphics g,
            XmlNode node,
            XmlDocument document,
            GraphicsPath path,
            float sx,
            float sy)
        {
            if (g == null ||
                node == null ||
                path == null ||
                path.PointCount == 0)
            {
                return;
            }

            float offsetX;
            float offsetY;
            float blurX;
            float blurY;
            Color shadowColor;

            if (!TryReadSvgDropShadow(
                    node,
                    document,
                    sx,
                    sy,
                    out offsetX,
                    out offsetY,
                    out blurX,
                    out blurY,
                    out shadowColor))
            {
                return;
            }

            string fillValue =
                GetSvgStyleInherited(
                    node,
                    "fill");

            bool hasFill =
                !string.Equals(
                    fillValue,
                    "none",
                    StringComparison.OrdinalIgnoreCase);

            Color stroke =
                ReadSvgColorInherited(
                    node,
                    "stroke",
                    Color.Transparent);

            bool hasStroke =
                stroke.A > 0;

            float strokeWidth =
                Math.Max(
                    1f,
                    ParseSvgFloat(
                        GetSvgStyleInherited(
                            node,
                            "stroke-width"),
                        1f) *
                    Math.Min(
                        Math.Abs(sx),
                        Math.Abs(sy)));

            float elementOpacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "opacity"),
                    1f);

            int sourceAlpha =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            shadowColor.A *
                            elementOpacity)));

            if (sourceAlpha <= 0)
            {
                return;
            }

            if (blurX <= 0.05f &&
                blurY <= 0.05f)
            {
                using (GraphicsPath shifted =
                    (GraphicsPath)path.Clone())
                using (Matrix translation =
                    new Matrix())
                {
                    translation.Translate(
                        offsetX,
                        offsetY);
                    shifted.Transform(
                        translation);

                    using (Brush brush =
                        new SolidBrush(
                            Color.FromArgb(
                                sourceAlpha,
                                shadowColor)))
                    {
                        if (hasFill)
                        {
                            g.FillPath(
                                brush,
                                shifted);
                        }
                    }

                    if (hasStroke)
                    {
                        using (Pen pen =
                            new Pen(
                                Color.FromArgb(
                                    sourceAlpha,
                                    shadowColor),
                                strokeWidth))
                        {
                            g.DrawPath(
                                pen,
                                shifted);
                        }
                    }
                }

                return;
            }

            const int rings = 4;
            const int directions = 12;

            for (int ring = rings;
                 ring >= 0;
                 ring--)
            {
                float ringRatio =
                    ring /
                    (float)rings;

                int directionCount =
                    ring == 0
                        ? 1
                        : directions;

                float alphaRatio =
                    ring == 0
                        ? 0.22f
                        : (rings -
                           ring +
                           1) /
                          (float)(rings *
                                  10);

                for (int direction = 0;
                     direction < directionCount;
                     direction++)
                {
                    double angle =
                        direction *
                        Math.PI *
                        2.0 /
                        Math.Max(
                            1,
                            directionCount);

                    float shadowX =
                        offsetX +
                        (float)Math.Cos(
                            angle) *
                        blurX *
                        ringRatio;

                    float shadowY =
                        offsetY +
                        (float)Math.Sin(
                            angle) *
                        blurY *
                        ringRatio;

                    int alpha =
                        Math.Max(
                            1,
                            Math.Min(
                                255,
                                (int)Math.Round(
                                    sourceAlpha *
                                    alphaRatio)));

                    using (GraphicsPath shifted =
                        (GraphicsPath)path.Clone())
                    using (Matrix translation =
                        new Matrix())
                    {
                        translation.Translate(
                            shadowX,
                            shadowY);
                        shifted.Transform(
                            translation);

                        if (hasFill)
                        {
                            using (Brush brush =
                                new SolidBrush(
                                    Color.FromArgb(
                                        alpha,
                                        shadowColor)))
                            {
                                g.FillPath(
                                    brush,
                                    shifted);
                            }
                        }

                        if (hasStroke)
                        {
                            using (Pen pen =
                                new Pen(
                                    Color.FromArgb(
                                        alpha,
                                        shadowColor),
                                    strokeWidth))
                            {
                                g.DrawPath(
                                    pen,
                                    shifted);
                            }
                        }
                    }
                }
            }
        }

        private static bool TryReadSvgGaussianBlur(
            XmlNode node,
            XmlDocument document,
            float sx,
            float sy,
            out float blurX,
            out float blurY)
        {
            blurX = 0f;
            blurY = 0f;

            if (node == null ||
                document == null)
            {
                return false;
            }

            string filterValue =
                GetSvgStyleInherited(
                    node,
                    "filter");

            string filterId =
                ExtractSvgUrlId(
                    filterValue);

            if (string.IsNullOrEmpty(
                    filterId))
            {
                return false;
            }

            XmlNode filter =
                FindSvgNodeById(
                    document,
                    filterId);

            if (filter == null ||
                filter.LocalName !=
                    "filter")
            {
                return false;
            }

            XmlNode gaussian =
                null;

            for (int i = 0;
                 i < filter.ChildNodes.Count;
                 i++)
            {
                XmlNode child =
                    filter.ChildNodes[i];

                if (child != null &&
                    child.LocalName ==
                        "feGaussianBlur")
                {
                    gaussian =
                        child;
                    break;
                }
            }

            if (gaussian == null)
            {
                return false;
            }

            string input =
                GetAttr(
                    gaussian,
                    "in");

            if (!string.IsNullOrEmpty(
                    input) &&
                !string.Equals(
                    input,
                    "SourceGraphic",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string raw =
                GetAttr(
                    gaussian,
                    "stdDeviation");

            if (string.IsNullOrEmpty(
                    raw))
            {
                return false;
            }

            string[] parts =
                raw.Replace(
                    ',',
                    ' ')
                .Split(
                    new char[]
                    {
                        ' ',
                        '\t',
                        '\r',
                        '\n'
                    },
                    StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
            {
                return false;
            }

            float stdX =
                Math.Max(
                    0f,
                    ParseSvgFloat(
                        parts[0],
                        0f));

            float stdY =
                parts.Length > 1
                    ? Math.Max(
                        0f,
                        ParseSvgFloat(
                            parts[1],
                            stdX))
                    : stdX;

            blurX =
                Math.Min(
                    48f,
                    stdX *
                    Math.Abs(sx));

            blurY =
                Math.Min(
                    48f,
                    stdY *
                    Math.Abs(sy));

            return blurX > 0.05f ||
                blurY > 0.05f;
        }

        private static void DrawSvgGaussianBlurApproximation(
            Graphics g,
            XmlNode node,
            XmlDocument document,
            GraphicsPath path,
            float sx,
            float sy)
        {
            if (g == null ||
                node == null ||
                path == null ||
                path.PointCount == 0)
            {
                return;
            }

            float blurX;
            float blurY;

            if (!TryReadSvgGaussianBlur(
                    node,
                    document,
                    sx,
                    sy,
                    out blurX,
                    out blurY))
            {
                return;
            }

            string fillValue =
                GetSvgStyleInherited(
                    node,
                    "fill");

            bool solidFill =
                !string.Equals(
                    fillValue,
                    "none",
                    StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrEmpty(
                    ExtractSvgUrlId(
                        fillValue));

            Color fill =
                ReadSvgColorInherited(
                    node,
                    "fill",
                    Color.Black);

            float overallOpacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "opacity"),
                    1f);

            float fillOpacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "fill-opacity"),
                    1f);

            int fillAlpha =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            fill.A *
                            overallOpacity *
                            fillOpacity)));

            Color stroke =
                ReadSvgColorInherited(
                    node,
                    "stroke",
                    Color.Transparent);

            float strokeOpacity =
                ReadSvgOpacity(
                    GetSvgStyleInherited(
                        node,
                        "stroke-opacity"),
                    1f);

            int strokeAlpha =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            stroke.A *
                            overallOpacity *
                            strokeOpacity)));

            float strokeWidth =
                Math.Max(
                    1f,
                    ParseSvgFloat(
                        GetSvgStyleInherited(
                            node,
                            "stroke-width"),
                        1f) *
                    Math.Min(
                        Math.Abs(sx),
                        Math.Abs(sy)));

            const int rings = 6;
            const int directions = 16;

            for (int ring = rings;
                 ring >= 1;
                 ring--)
            {
                float ringRatio =
                    ring /
                    (float)rings;

                double sigmaDistance =
                    3.0 *
                    ringRatio;

                float gaussianWeight =
                    (float)Math.Exp(
                        -0.5 *
                        sigmaDistance *
                        sigmaDistance);

                float alphaRatio =
                    Math.Max(
                        0.006f,
                        gaussianWeight /
                        (directions *
                         0.65f));

                for (int direction = 0;
                     direction < directions;
                     direction++)
                {
                    double angle =
                        direction *
                        Math.PI *
                        2.0 /
                        directions;

                    float offsetX =
                        (float)Math.Cos(
                            angle) *
                        blurX *
                        3f *
                        ringRatio;
                    float offsetY =
                        (float)Math.Sin(
                            angle) *
                        blurY *
                        3f *
                        ringRatio;

                    using (GraphicsPath shifted =
                        (GraphicsPath)path.Clone())
                    using (Matrix translation =
                        new Matrix())
                    {
                        translation.Translate(
                            offsetX,
                            offsetY);
                        shifted.Transform(
                            translation);

                        if (solidFill &&
                            fillAlpha > 0)
                        {
                            int alpha =
                                Math.Max(
                                    1,
                                    Math.Min(
                                        255,
                                        (int)Math.Round(
                                            fillAlpha *
                                            alphaRatio)));

                            using (Brush brush =
                                new SolidBrush(
                                    Color.FromArgb(
                                        alpha,
                                        fill)))
                            {
                                g.FillPath(
                                    brush,
                                    shifted);
                            }
                        }

                        if (strokeAlpha > 0)
                        {
                            int alpha =
                                Math.Max(
                                    1,
                                    Math.Min(
                                        255,
                                        (int)Math.Round(
                                            strokeAlpha *
                                            alphaRatio)));

                            using (Pen pen =
                                new Pen(
                                    Color.FromArgb(
                                        alpha,
                                        stroke),
                                    strokeWidth))
                            {
                                g.DrawPath(
                                    pen,
                                    shifted);
                            }
                        }
                    }
                }
            }
        }

        private static XmlNode FindSvgNodeById(
            XmlDocument document,
            string id)
        {
            if (document == null ||
                document.DocumentElement == null ||
                string.IsNullOrEmpty(id))
            {
                return null;
            }

            return FindSvgNodeById(
                document.DocumentElement,
                id);
        }

        private static XmlNode FindSvgNodeById(
            XmlNode node,
            string id)
        {
            if (node == null)
                return null;

            if (string.Equals(
                    GetAttr(
                        node,
                        "id"),
                    id,
                    StringComparison.Ordinal))
            {
                return node;
            }

            for (int i = 0;
                 i < node.ChildNodes.Count;
                 i++)
            {
                XmlNode found =
                    FindSvgNodeById(
                        node.ChildNodes[i],
                        id);

                if (found != null)
                    return found;
            }

            return null;
        }

        private static string ExtractSvgUrlId(
            string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return null;

            raw = raw.Trim();

            if (!raw.StartsWith(
                    "url(",
                    StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            int hash =
                raw.IndexOf('#');
            int close =
                raw.IndexOf(
                    ')',
                    Math.Max(
                        0,
                        hash));

            if (hash < 0 ||
                close <= hash + 1)
            {
                return null;
            }

            return raw
                .Substring(
                    hash + 1,
                    close -
                    hash -
                    1)
                .Trim(
                    ' ',
                    '\'',
                    '"');
        }

        private static string GetSvgStyleInherited(
            XmlNode node,
            string key)
        {
            XmlNode current =
                node;

            while (current != null)
            {
                string value =
                    GetSvgStyle(
                        current,
                        key);

                if (!string.IsNullOrEmpty(
                        value))
                {
                    return value;
                }

                if (current.LocalName ==
                    "svg")
                {
                    break;
                }

                current =
                    current.ParentNode;
            }

            return null;
        }

        private static Color ReadSvgColorInherited(
            XmlNode node,
            string key,
            Color fallback)
        {
            string raw =
                GetSvgStyleInherited(
                    node,
                    key);

            if (string.IsNullOrEmpty(raw))
                return fallback;

            raw = raw.Trim();

            if (raw == "none")
                return Color.Transparent;

            if (raw.StartsWith("#"))
            {
                Color? parsed =
                    ParseHexColor(
                        raw.TrimStart('#'));

                if (parsed.HasValue)
                    return parsed.Value;
            }

            Color named =
                Color.FromName(raw);

            return named.A > 0 ||
                string.Equals(
                    raw,
                    "transparent",
                    StringComparison.OrdinalIgnoreCase)
                    ? named
                    : fallback;
        }

        private static float ReadSvgOpacity(
            string raw,
            float fallback)
        {
            if (string.IsNullOrEmpty(raw))
                return fallback;

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                float percent =
                    ParseSvgFloat(
                        raw.Substring(
                            0,
                            raw.Length - 1),
                        fallback *
                        100f);

                return Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        percent /
                        100f));
            }

            float value =
                ParseSvgFloat(
                    raw,
                    fallback);

            return Math.Max(
                0f,
                Math.Min(
                    1f,
                    value));
        }

        private static float ReadSvgGradientOffset(
            string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return 0f;

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                return Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        ParseSvgFloat(
                            raw.Substring(
                                0,
                                raw.Length - 1),
                            0f) /
                        100f));
            }

            return Math.Max(
                0f,
                Math.Min(
                    1f,
                    ParseSvgFloat(
                        raw,
                        0f)));
        }

        private static float ResolveSvgGradientCoordinate(
            string raw,
            float boundsStart,
            float boundsLength,
            float targetStart,
            float viewMin,
            float scale,
            float fallbackFraction,
            bool userSpace)
        {
            if (string.IsNullOrEmpty(raw))
            {
                return boundsStart +
                    boundsLength *
                    fallbackFraction;
            }

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                float percentage =
                    ParseSvgFloat(
                        raw.Substring(
                            0,
                            raw.Length - 1),
                        fallbackFraction *
                        100f) /
                    100f;

                return boundsStart +
                    boundsLength *
                    percentage;
            }

            float value =
                ParseSvgFloat(
                    raw,
                    fallbackFraction);

            if (userSpace)
            {
                return targetStart +
                    (value -
                     viewMin) *
                    scale;
            }

            return boundsStart +
                boundsLength *
                value;
        }

        private static float ResolveSvgGradientRadius(
            string raw,
            RectangleF bounds,
            float scale,
            float fallbackFraction,
            bool userSpace)
        {
            float reference =
                Math.Min(
                    bounds.Width,
                    bounds.Height);

            if (string.IsNullOrEmpty(raw))
            {
                return reference *
                    fallbackFraction;
            }

            raw = raw.Trim();

            if (raw.EndsWith("%"))
            {
                return reference *
                    ParseSvgFloat(
                        raw.Substring(
                            0,
                            raw.Length - 1),
                        fallbackFraction *
                        100f) /
                    100f;
            }

            float value =
                ParseSvgFloat(
                    raw,
                    fallbackFraction);

            return userSpace
                ? Math.Abs(
                    value *
                    scale)
                : reference *
                    value;
        }

        private static void ApplySimpleSvgTransform(
            Graphics g,
            XmlNode node,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            string transform =
                GetAttr(
                    node,
                    "transform");

            if (string.IsNullOrEmpty(transform))
                return;

            ApplySvgTransformSequence(
                g,
                transform,
                target,
                minX,
                minY,
                sx,
                sy);
        }

        private static void ApplySvgTransformSequence(
            Graphics g,
            string transform,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            if (g == null ||
                string.IsNullOrEmpty(transform))
            {
                return;
            }

            int index = 0;

            while (index < transform.Length)
            {
                while (index < transform.Length &&
                    (char.IsWhiteSpace(transform[index]) ||
                     transform[index] == ','))
                {
                    index++;
                }

                int nameStart = index;

                while (index < transform.Length &&
                    char.IsLetter(transform[index]))
                {
                    index++;
                }

                if (index <= nameStart)
                {
                    index++;
                    continue;
                }

                string name =
                    transform.Substring(
                        nameStart,
                        index -
                        nameStart);

                while (index < transform.Length &&
                    char.IsWhiteSpace(transform[index]))
                {
                    index++;
                }

                if (index >= transform.Length ||
                    transform[index] != '(')
                {
                    continue;
                }

                int argumentStart =
                    index + 1;
                int close =
                    transform.IndexOf(
                        ')',
                        argumentStart);

                if (close < 0)
                    break;

                List<float> values =
                    ParseSvgNumberList(
                        transform.Substring(
                            argumentStart,
                            close -
                            argumentStart));

                ApplySvgTransformOperation(
                    g,
                    name,
                    values,
                    target,
                    minX,
                    minY,
                    sx,
                    sy);

                index =
                    close + 1;
            }
        }

        private static void ApplySvgTransformOperation(
            Graphics g,
            string name,
            List<float> values,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            if (g == null ||
                string.IsNullOrEmpty(name) ||
                values == null)
            {
                return;
            }

            string lower =
                name.ToLowerInvariant();

            if (lower == "matrix" &&
                values.Count >= 6)
            {
                float safeSx =
                    Math.Abs(sx) < 0.0001f
                        ? 1f
                        : sx;
                float safeSy =
                    Math.Abs(sy) < 0.0001f
                        ? 1f
                        : sy;

                using (Matrix matrix =
                    new Matrix(
                        values[0],
                        values[1] *
                            safeSy /
                            safeSx,
                        values[2] *
                            safeSx /
                            safeSy,
                        values[3],
                        values[4] *
                            safeSx,
                        values[5] *
                            safeSy))
                {
                    g.MultiplyTransform(
                        matrix,
                        MatrixOrder.Append);
                }
                return;
            }

            if (lower == "translate" &&
                values.Count > 0)
            {
                float dx =
                    values[0] *
                    sx;
                float dy =
                    values.Count > 1
                        ? values[1] *
                            sy
                        : 0f;

                g.TranslateTransform(
                    dx,
                    dy,
                    MatrixOrder.Append);
                return;
            }

            if (lower == "scale" &&
                values.Count > 0)
            {
                float scaleX =
                    values[0];
                float scaleY =
                    values.Count > 1
                        ? values[1]
                        : scaleX;

                g.ScaleTransform(
                    scaleX,
                    scaleY,
                    MatrixOrder.Append);
                return;
            }

            if (lower == "rotate" &&
                values.Count > 0)
            {
                float degrees =
                    values[0];

                if (values.Count >= 3)
                {
                    float cx =
                        SvgX(
                            target,
                            minX,
                            sx,
                            values[1]);

                    float cy =
                        SvgY(
                            target,
                            minY,
                            sy,
                            values[2]);

                    g.TranslateTransform(
                        cx,
                        cy,
                        MatrixOrder.Append);
                    g.RotateTransform(
                        degrees,
                        MatrixOrder.Append);
                    g.TranslateTransform(
                        -cx,
                        -cy,
                        MatrixOrder.Append);
                }
                else
                {
                    g.RotateTransform(
                        degrees,
                        MatrixOrder.Append);
                }
                return;
            }

            if (lower == "skewx" &&
                values.Count > 0)
            {
                float tangent =
                    (float)Math.Tan(
                        values[0] *
                        Math.PI /
                        180.0);

                using (Matrix matrix =
                    new Matrix(
                        1f,
                        0f,
                        tangent,
                        1f,
                        0f,
                        0f))
                {
                    g.MultiplyTransform(
                        matrix,
                        MatrixOrder.Append);
                }
                return;
            }

            if (lower == "skewy" &&
                values.Count > 0)
            {
                float tangent =
                    (float)Math.Tan(
                        values[0] *
                        Math.PI /
                        180.0);

                using (Matrix matrix =
                    new Matrix(
                        1f,
                        tangent,
                        0f,
                        1f,
                        0f,
                        0f))
                {
                    g.MultiplyTransform(
                        matrix,
                        MatrixOrder.Append);
                }
            }
        }

        private static void ApplySvgGradientSpreadMethod(
            Brush brush,
            XmlNode gradient)
        {
            if (brush == null ||
                gradient == null)
            {
                return;
            }

            string spread =
                GetAttr(
                    gradient,
                    "spreadMethod");

            if (string.IsNullOrEmpty(spread) ||
                string.Equals(
                    spread,
                    "pad",
                    StringComparison.OrdinalIgnoreCase))
            {
                LinearGradientBrush linearClamp =
                    brush as LinearGradientBrush;

                if (linearClamp != null)
                {
                    linearClamp.WrapMode =
                        WrapMode.Clamp;
                }

                PathGradientBrush radialClamp =
                    brush as PathGradientBrush;

                if (radialClamp != null)
                {
                    radialClamp.WrapMode =
                        WrapMode.Clamp;
                }

                return;
            }

            WrapMode mode =
                string.Equals(
                    spread,
                    "reflect",
                    StringComparison.OrdinalIgnoreCase)
                    ? WrapMode.TileFlipXY
                    : WrapMode.Tile;

            LinearGradientBrush linear =
                brush as LinearGradientBrush;

            if (linear != null)
            {
                linear.WrapMode = mode;
                return;
            }

            PathGradientBrush radial =
                brush as PathGradientBrush;

            if (radial != null)
            {
                radial.WrapMode = mode;
            }
        }

        private static void ApplySvgGradientTransform(
            Brush brush,
            XmlNode gradient,
            RectangleF bounds,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy,
            bool userSpace)
        {
            if (brush == null ||
                gradient == null)
            {
                return;
            }

            string transform =
                GetAttr(
                    gradient,
                    "gradientTransform");

            if (string.IsNullOrEmpty(transform))
                return;

            Matrix matrix =
                BuildSvgBrushTransformMatrix(
                    transform,
                    bounds,
                    target,
                    minX,
                    minY,
                    sx,
                    sy,
                    userSpace);

            if (matrix == null)
                return;

            try
            {
                LinearGradientBrush linear =
                    brush as LinearGradientBrush;

                if (linear != null)
                {
                    linear.MultiplyTransform(
                        matrix,
                        MatrixOrder.Append);
                    return;
                }

                PathGradientBrush radial =
                    brush as PathGradientBrush;

                if (radial != null)
                {
                    radial.MultiplyTransform(
                        matrix,
                        MatrixOrder.Append);
                }
            }
            finally
            {
                matrix.Dispose();
            }
        }

        private static Matrix BuildSvgBrushTransformMatrix(
            string transform,
            RectangleF bounds,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy,
            bool userSpace)
        {
            Matrix result =
                new Matrix();

            int index = 0;

            while (index < transform.Length)
            {
                while (index < transform.Length &&
                    (char.IsWhiteSpace(transform[index]) ||
                     transform[index] == ','))
                {
                    index++;
                }

                int nameStart = index;

                while (index < transform.Length &&
                    char.IsLetter(transform[index]))
                {
                    index++;
                }

                if (index <= nameStart)
                {
                    index++;
                    continue;
                }

                string name =
                    transform.Substring(
                        nameStart,
                        index -
                        nameStart)
                    .ToLowerInvariant();

                while (index < transform.Length &&
                    char.IsWhiteSpace(transform[index]))
                {
                    index++;
                }

                if (index >= transform.Length ||
                    transform[index] != '(')
                {
                    continue;
                }

                int argumentStart =
                    index + 1;
                int close =
                    transform.IndexOf(
                        ')',
                        argumentStart);

                if (close < 0)
                    break;

                List<float> values =
                    ParseSvgNumberList(
                        transform.Substring(
                            argumentStart,
                            close -
                            argumentStart));

                float coordinateSx =
                    userSpace
                        ? sx
                        : bounds.Width;
                float coordinateSy =
                    userSpace
                        ? sy
                        : bounds.Height;

                if (name == "matrix" &&
                    values.Count >= 6)
                {
                    float safeX =
                        Math.Abs(coordinateSx) <
                        0.0001f
                            ? 1f
                            : coordinateSx;
                    float safeY =
                        Math.Abs(coordinateSy) <
                        0.0001f
                            ? 1f
                            : coordinateSy;

                    using (Matrix operation =
                        new Matrix(
                            values[0],
                            values[1] *
                                safeY /
                                safeX,
                            values[2] *
                                safeX /
                                safeY,
                            values[3],
                            values[4] *
                                coordinateSx,
                            values[5] *
                                coordinateSy))
                    {
                        result.Multiply(
                            operation,
                            MatrixOrder.Append);
                    }
                }
                else if (name == "translate" &&
                    values.Count > 0)
                {
                    result.Translate(
                        values[0] *
                            coordinateSx,
                        (values.Count > 1
                            ? values[1]
                            : 0f) *
                            coordinateSy,
                        MatrixOrder.Append);
                }
                else if (name == "scale" &&
                    values.Count > 0)
                {
                    result.Scale(
                        values[0],
                        values.Count > 1
                            ? values[1]
                            : values[0],
                        MatrixOrder.Append);
                }
                else if (name == "rotate" &&
                    values.Count > 0)
                {
                    if (values.Count >= 3)
                    {
                        float cx =
                            userSpace
                                ? SvgX(
                                    target,
                                    minX,
                                    sx,
                                    values[1])
                                : bounds.Left +
                                    bounds.Width *
                                    values[1];

                        float cy =
                            userSpace
                                ? SvgY(
                                    target,
                                    minY,
                                    sy,
                                    values[2])
                                : bounds.Top +
                                    bounds.Height *
                                    values[2];

                        result.Translate(
                            cx,
                            cy,
                            MatrixOrder.Append);
                        result.Rotate(
                            values[0],
                            MatrixOrder.Append);
                        result.Translate(
                            -cx,
                            -cy,
                            MatrixOrder.Append);
                    }
                    else
                    {
                        result.RotateAt(
                            values[0],
                            new PointF(
                                bounds.Left +
                                    bounds.Width /
                                    2f,
                                bounds.Top +
                                    bounds.Height /
                                    2f),
                            MatrixOrder.Append);
                    }
                }

                index =
                    close + 1;
            }

            return result;
        }

        private static string ReadSvgTransformArguments(
            string raw,
            int start)
        {
            if (string.IsNullOrEmpty(raw) ||
                start < 0 ||
                start >= raw.Length)
            {
                return string.Empty;
            }

            int end =
                raw.IndexOf(
                    ')',
                    start);

            if (end < 0)
                end = raw.Length;

            return raw.Substring(
                start,
                end - start);
        }

        private static List<float> ParseSvgNumberList(
            string raw)
        {
            List<float> result =
                new List<float>();

            if (string.IsNullOrEmpty(raw))
                return result;

            string cleaned =
                raw.Replace(
                    ',',
                    ' ');

            string[] pieces =
                cleaned.Split(
                    new char[]
                    {
                        ' ',
                        '\t',
                        '\r',
                        '\n'
                    },
                    StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0;
                 i < pieces.Length;
                 i++)
            {
                float value;
                if (float.TryParse(
                        pieces[i],
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out value))
                {
                    result.Add(
                        value);
                }
            }

            return result;
        }

        private static bool TryDrawSimpleSvgFromDocument(
            Graphics g,
            XmlDocument document,
            RectangleF target)
        {
            if (document == null || document.DocumentElement == null)
                return false;

            XmlNode svg = document.DocumentElement;
            float minX = 0f;
            float minY = 0f;
            float width = ParseSvgFloat(GetAttr(svg, "width"), 100f);
            float height = ParseSvgFloat(GetAttr(svg, "height"), 100f);
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
                    width = Math.Max(0.001f, ParseSvgFloat(pieces[2], width));
                    height = Math.Max(0.001f, ParseSvgFloat(pieces[3], height));
                }
            }

            DrawSvgChildren(
                g,
                svg,
                target,
                minX,
                minY,
                target.Width / Math.Max(0.001f, width),
                target.Height / Math.Max(0.001f, height));
            return true;
        }

        private static GraphicsPath BuildSvgPath(
            string data,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            List<string> tokens =
                TokenizeSvgPath(data);

            if (tokens.Count == 0)
                return null;

            GraphicsPath path =
                new GraphicsPath();

            int index = 0;
            char command = ' ';
            char previousUpper = ' ';
            PointF current =
                new PointF(
                    0f,
                    0f);
            PointF start =
                current;
            PointF lastCubicControl =
                current;
            PointF lastQuadraticControl =
                current;
            bool hasCubicControl = false;
            bool hasQuadraticControl = false;

            while (index < tokens.Count)
            {
                if (IsSvgCommand(
                        tokens[index]))
                {
                    command =
                        tokens[index][0];
                    index++;

                    if (command == 'Z' ||
                        command == 'z')
                    {
                        path.CloseFigure();
                        current = start;
                        previousUpper = 'Z';
                        hasCubicControl = false;
                        hasQuadraticControl = false;
                        continue;
                    }
                }

                bool relative =
                    char.IsLower(
                        command);

                char upper =
                    char.ToUpperInvariant(
                        command);

                float a, b, cc, d, e, f, h;

                if (upper == 'M' ||
                    upper == 'L')
                {
                    if (!ReadSvgNumber(
                            tokens,
                            ref index,
                            out a) ||
                        !ReadSvgNumber(
                            tokens,
                            ref index,
                            out b))
                    {
                        break;
                    }

                    PointF next =
                        SvgPoint(
                            a,
                            b,
                            relative,
                            current);

                    if (upper == 'M')
                    {
                        path.StartFigure();
                        start = next;
                        current = next;
                        previousUpper = 'M';
                        command =
                            relative
                                ? 'l'
                                : 'L';
                    }
                    else
                    {
                        path.AddLine(
                            ToSvgTarget(
                                current,
                                target,
                                minX,
                                minY,
                                sx,
                                sy),
                            ToSvgTarget(
                                next,
                                target,
                                minX,
                                minY,
                                sx,
                                sy));
                        current = next;
                        previousUpper = 'L';
                    }

                    hasCubicControl = false;
                    hasQuadraticControl = false;
                }
                else if (upper == 'H')
                {
                    if (!ReadSvgNumber(
                            tokens,
                            ref index,
                            out a))
                    {
                        break;
                    }

                    PointF next =
                        new PointF(
                            relative
                                ? current.X + a
                                : a,
                            current.Y);

                    path.AddLine(
                        ToSvgTarget(
                            current,
                            target,
                            minX,
                            minY,
                            sx,
                            sy),
                        ToSvgTarget(
                            next,
                            target,
                            minX,
                            minY,
                            sx,
                            sy));

                    current = next;
                    previousUpper = 'H';
                    hasCubicControl = false;
                    hasQuadraticControl = false;
                }
                else if (upper == 'V')
                {
                    if (!ReadSvgNumber(
                            tokens,
                            ref index,
                            out a))
                    {
                        break;
                    }

                    PointF next =
                        new PointF(
                            current.X,
                            relative
                                ? current.Y + a
                                : a);

                    path.AddLine(
                        ToSvgTarget(
                            current,
                            target,
                            minX,
                            minY,
                            sx,
                            sy),
                        ToSvgTarget(
                            next,
                            target,
                            minX,
                            minY,
                            sx,
                            sy));

                    current = next;
                    previousUpper = 'V';
                    hasCubicControl = false;
                    hasQuadraticControl = false;
                }
                else if (upper == 'C')
                {
                    if (!ReadSvgNumber(tokens, ref index, out a) ||
                        !ReadSvgNumber(tokens, ref index, out b) ||
                        !ReadSvgNumber(tokens, ref index, out cc) ||
                        !ReadSvgNumber(tokens, ref index, out d) ||
                        !ReadSvgNumber(tokens, ref index, out e) ||
                        !ReadSvgNumber(tokens, ref index, out f))
                    {
                        break;
                    }

                    PointF c1 =
                        SvgPoint(
                            a,
                            b,
                            relative,
                            current);
                    PointF c2 =
                        SvgPoint(
                            cc,
                            d,
                            relative,
                            current);
                    PointF finish =
                        SvgPoint(
                            e,
                            f,
                            relative,
                            current);

                    path.AddBezier(
                        ToSvgTarget(current, target, minX, minY, sx, sy),
                        ToSvgTarget(c1, target, minX, minY, sx, sy),
                        ToSvgTarget(c2, target, minX, minY, sx, sy),
                        ToSvgTarget(finish, target, minX, minY, sx, sy));

                    current = finish;
                    lastCubicControl = c2;
                    hasCubicControl = true;
                    hasQuadraticControl = false;
                    previousUpper = 'C';
                }
                else if (upper == 'S')
                {
                    if (!ReadSvgNumber(tokens, ref index, out a) ||
                        !ReadSvgNumber(tokens, ref index, out b) ||
                        !ReadSvgNumber(tokens, ref index, out cc) ||
                        !ReadSvgNumber(tokens, ref index, out d))
                    {
                        break;
                    }

                    PointF c1 =
                        hasCubicControl &&
                        (previousUpper == 'C' ||
                         previousUpper == 'S')
                            ? new PointF(
                                current.X * 2f -
                                    lastCubicControl.X,
                                current.Y * 2f -
                                    lastCubicControl.Y)
                            : current;

                    PointF c2 =
                        SvgPoint(
                            a,
                            b,
                            relative,
                            current);

                    PointF finish =
                        SvgPoint(
                            cc,
                            d,
                            relative,
                            current);

                    path.AddBezier(
                        ToSvgTarget(current, target, minX, minY, sx, sy),
                        ToSvgTarget(c1, target, minX, minY, sx, sy),
                        ToSvgTarget(c2, target, minX, minY, sx, sy),
                        ToSvgTarget(finish, target, minX, minY, sx, sy));

                    current = finish;
                    lastCubicControl = c2;
                    hasCubicControl = true;
                    hasQuadraticControl = false;
                    previousUpper = 'S';
                }
                else if (upper == 'Q')
                {
                    if (!ReadSvgNumber(tokens, ref index, out a) ||
                        !ReadSvgNumber(tokens, ref index, out b) ||
                        !ReadSvgNumber(tokens, ref index, out cc) ||
                        !ReadSvgNumber(tokens, ref index, out d))
                    {
                        break;
                    }

                    PointF control =
                        SvgPoint(
                            a,
                            b,
                            relative,
                            current);
                    PointF finish =
                        SvgPoint(
                            cc,
                            d,
                            relative,
                            current);

                    AppendSvgQuadratic(
                        path,
                        current,
                        control,
                        finish,
                        target,
                        minX,
                        minY,
                        sx,
                        sy);

                    current = finish;
                    lastQuadraticControl = control;
                    hasQuadraticControl = true;
                    hasCubicControl = false;
                    previousUpper = 'Q';
                }
                else if (upper == 'T')
                {
                    if (!ReadSvgNumber(tokens, ref index, out a) ||
                        !ReadSvgNumber(tokens, ref index, out b))
                    {
                        break;
                    }

                    PointF control =
                        hasQuadraticControl &&
                        (previousUpper == 'Q' ||
                         previousUpper == 'T')
                            ? new PointF(
                                current.X * 2f -
                                    lastQuadraticControl.X,
                                current.Y * 2f -
                                    lastQuadraticControl.Y)
                            : current;

                    PointF finish =
                        SvgPoint(
                            a,
                            b,
                            relative,
                            current);

                    AppendSvgQuadratic(
                        path,
                        current,
                        control,
                        finish,
                        target,
                        minX,
                        minY,
                        sx,
                        sy);

                    current = finish;
                    lastQuadraticControl = control;
                    hasQuadraticControl = true;
                    hasCubicControl = false;
                    previousUpper = 'T';
                }
                else if (upper == 'A')
                {
                    if (!ReadSvgNumber(tokens, ref index, out a) ||
                        !ReadSvgNumber(tokens, ref index, out b) ||
                        !ReadSvgNumber(tokens, ref index, out cc) ||
                        !ReadSvgNumber(tokens, ref index, out d) ||
                        !ReadSvgNumber(tokens, ref index, out e) ||
                        !ReadSvgNumber(tokens, ref index, out f) ||
                        !ReadSvgNumber(tokens, ref index, out h))
                    {
                        break;
                    }

                    PointF finish =
                        SvgPoint(
                            f,
                            h,
                            relative,
                            current);

                    AppendSvgArc(
                        path,
                        current,
                        finish,
                        a,
                        b,
                        cc,
                        Math.Abs(d) > 0.5f,
                        Math.Abs(e) > 0.5f,
                        target,
                        minX,
                        minY,
                        sx,
                        sy);

                    current = finish;
                    hasCubicControl = false;
                    hasQuadraticControl = false;
                    previousUpper = 'A';
                }
                else
                {
                    index++;
                    hasCubicControl = false;
                    hasQuadraticControl = false;
                    previousUpper = upper;
                }
            }

            return path;
        }

        private static void AppendSvgQuadratic(
            GraphicsPath path,
            PointF start,
            PointF control,
            PointF finish,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            PointF c1 =
                new PointF(
                    start.X +
                        (control.X -
                         start.X) *
                        2f /
                        3f,
                    start.Y +
                        (control.Y -
                         start.Y) *
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

            path.AddBezier(
                ToSvgTarget(start, target, minX, minY, sx, sy),
                ToSvgTarget(c1, target, minX, minY, sx, sy),
                ToSvgTarget(c2, target, minX, minY, sx, sy),
                ToSvgTarget(finish, target, minX, minY, sx, sy));
        }

        private static void AppendSvgArc(
            GraphicsPath path,
            PointF start,
            PointF end,
            float radiusX,
            float radiusY,
            float rotationDegrees,
            bool largeArc,
            bool sweep,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            double rx =
                Math.Abs(
                    radiusX);
            double ry =
                Math.Abs(
                    radiusY);

            if (rx < 0.000001 ||
                ry < 0.000001 ||
                (Math.Abs(
                    start.X -
                    end.X) < 0.000001 &&
                 Math.Abs(
                    start.Y -
                    end.Y) < 0.000001))
            {
                path.AddLine(
                    ToSvgTarget(
                        start,
                        target,
                        minX,
                        minY,
                        sx,
                        sy),
                    ToSvgTarget(
                        end,
                        target,
                        minX,
                        minY,
                        sx,
                        sy));
                return;
            }

            double phi =
                rotationDegrees *
                Math.PI /
                180.0;

            double cosPhi =
                Math.Cos(phi);
            double sinPhi =
                Math.Sin(phi);

            double dx =
                (start.X -
                 end.X) /
                2.0;
            double dy =
                (start.Y -
                 end.Y) /
                2.0;

            double x1p =
                cosPhi * dx +
                sinPhi * dy;
            double y1p =
                -sinPhi * dx +
                cosPhi * dy;

            double lambda =
                x1p * x1p /
                (rx * rx) +
                y1p * y1p /
                (ry * ry);

            if (lambda > 1.0)
            {
                double scale =
                    Math.Sqrt(lambda);
                rx *= scale;
                ry *= scale;
            }

            double rx2 = rx * rx;
            double ry2 = ry * ry;
            double x1p2 = x1p * x1p;
            double y1p2 = y1p * y1p;

            double numerator =
                rx2 * ry2 -
                rx2 * y1p2 -
                ry2 * x1p2;

            double denominator =
                rx2 * y1p2 +
                ry2 * x1p2;

            double factor = 0.0;
            if (denominator > 0.0000001)
            {
                factor =
                    Math.Sqrt(
                        Math.Max(
                            0.0,
                            numerator /
                            denominator));
            }

            if (largeArc == sweep)
                factor = -factor;

            double cxp =
                factor *
                rx *
                y1p /
                ry;
            double cyp =
                factor *
                -ry *
                x1p /
                rx;

            double cx =
                cosPhi * cxp -
                sinPhi * cyp +
                (start.X +
                 end.X) /
                2.0;
            double cy =
                sinPhi * cxp +
                cosPhi * cyp +
                (start.Y +
                 end.Y) /
                2.0;

            double ux =
                (x1p -
                 cxp) /
                rx;
            double uy =
                (y1p -
                 cyp) /
                ry;
            double vx =
                (-x1p -
                 cxp) /
                rx;
            double vy =
                (-y1p -
                 cyp) /
                ry;

            double startAngle =
                Math.Atan2(
                    uy,
                    ux);

            double deltaAngle =
                Math.Atan2(
                    ux * vy -
                    uy * vx,
                    ux * vx +
                    uy * vy);

            if (!sweep &&
                deltaAngle > 0.0)
            {
                deltaAngle -=
                    Math.PI *
                    2.0;
            }
            else if (sweep &&
                     deltaAngle < 0.0)
            {
                deltaAngle +=
                    Math.PI *
                    2.0;
            }

            int segments =
                Math.Max(
                    2,
                    Math.Min(
                        64,
                        (int)Math.Ceiling(
                            Math.Abs(
                                deltaAngle) /
                            (Math.PI /
                             12.0))));

            PointF previous =
                ToSvgTarget(
                    start,
                    target,
                    minX,
                    minY,
                    sx,
                    sy);

            for (int i = 1;
                 i <= segments;
                 i++)
            {
                double t =
                    startAngle +
                    deltaAngle *
                    i /
                    segments;

                double cosT =
                    Math.Cos(t);
                double sinT =
                    Math.Sin(t);

                double x =
                    cx +
                    cosPhi *
                    rx *
                    cosT -
                    sinPhi *
                    ry *
                    sinT;

                double y =
                    cy +
                    sinPhi *
                    rx *
                    cosT +
                    cosPhi *
                    ry *
                    sinT;

                PointF next =
                    ToSvgTarget(
                        new PointF(
                            (float)x,
                            (float)y),
                        target,
                        minX,
                        minY,
                        sx,
                        sy);

                path.AddLine(
                    previous,
                    next);
                previous = next;
            }
        }

        private static List<string> TokenizeSvgPath(string data)
        {
            List<string> tokens = new List<string>();
            if (string.IsNullOrEmpty(data))
                return tokens;

            int i = 0;
            while (i < data.Length)
            {
                char ch = data[i];
                if (char.IsLetter(ch))
                {
                    tokens.Add(ch.ToString());
                    i++;
                    continue;
                }

                if (char.IsWhiteSpace(ch) || ch == ',')
                {
                    i++;
                    continue;
                }

                int start = i;
                if (ch == '+' || ch == '-') i++;
                bool seenDot = false;
                bool seenExponent = false;

                while (i < data.Length)
                {
                    ch = data[i];
                    if (char.IsDigit(ch))
                    {
                        i++;
                        continue;
                    }
                    if (ch == '.' && !seenDot && !seenExponent)
                    {
                        seenDot = true;
                        i++;
                        continue;
                    }
                    if ((ch == 'e' || ch == 'E') && !seenExponent)
                    {
                        seenExponent = true;
                        i++;
                        if (i < data.Length && (data[i] == '+' || data[i] == '-')) i++;
                        continue;
                    }
                    break;
                }

                if (i > start)
                    tokens.Add(data.Substring(start, i - start));
                else
                    i++;
            }
            return tokens;
        }

        private static bool IsSvgCommand(string token)
        {
            return !string.IsNullOrEmpty(token) && token.Length == 1 && char.IsLetter(token[0]);
        }

        private static bool ReadSvgNumber(
            List<string> tokens,
            ref int index,
            out float value)
        {
            value = 0f;
            if (index >= tokens.Count || IsSvgCommand(tokens[index]))
                return false;
            bool ok = float.TryParse(
                tokens[index],
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out value);
            index++;
            return ok;
        }

        private static PointF SvgPoint(
            float x,
            float y,
            bool relative,
            PointF current)
        {
            return relative
                ? new PointF(current.X + x, current.Y + y)
                : new PointF(x, y);
        }

        private static PointF ToSvgTarget(
            PointF point,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            return new PointF(
                target.Left + (point.X - minX) * sx,
                target.Top + (point.Y - minY) * sy);
        }

        private static void DrawEnhancedGraphicFrame(
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
            string uri = graphicData == null ? string.Empty : GetAttr(graphicData, "uri");

            if (uri.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                XmlNode chart = FindFirst(frame, "chart");
                string rid = GetRelationshipId(chart);
                if (!string.IsNullOrEmpty(rid) && rels.ContainsKey(rid))
                {
                    XmlDocument chartDoc = LoadXml(zip, rels[rid]);
                    if (chartDoc != null && DrawExtendedChart(g, chartDoc, rect, theme))
                        return;
                }
            }

            if (uri.IndexOf("diagram", StringComparison.OrdinalIgnoreCase) >= 0 ||
                FindFirst(frame, "relIds") != null)
            {
                XmlNode relIds = FindFirst(frame, "relIds");
                string dataRid = null;
                string layoutRid = null;

                if (relIds != null && relIds.Attributes != null)
                {
                    for (int i = 0; i < relIds.Attributes.Count; i++)
                    {
                        XmlAttribute attribute = relIds.Attributes[i];

                        if (attribute.LocalName == "dm")
                            dataRid = attribute.Value;
                        else if (attribute.LocalName == "lo")
                            layoutRid = attribute.Value;
                    }
                }

                if (!string.IsNullOrEmpty(dataRid) &&
                    rels.ContainsKey(dataRid))
                {
                    XmlDocument dataDoc =
                        LoadXml(zip, rels[dataRid]);

                    XmlDocument layoutDoc = null;
                    if (!string.IsNullOrEmpty(layoutRid) &&
                        rels.ContainsKey(layoutRid))
                    {
                        layoutDoc =
                            LoadXml(zip, rels[layoutRid]);
                    }

                    if (dataDoc != null &&
                        DrawStructuredSmartArt(
                            g,
                            dataDoc,
                            layoutDoc,
                            rect,
                            theme))
                    {
                        return;
                    }
                }
            }

            DrawGraphicFrame(zip, g, frame, rels, ctx, theme);
        }

        private static bool DrawExtendedChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (chartDoc == null)
                return false;

            if (FindFirst(chartDoc, "doughnutChart") != null)
            {
                DrawDoughnutChart(g, chartDoc, rect, theme);
                return true;
            }
            if (FindFirst(chartDoc, "scatterChart") != null ||
                FindFirst(chartDoc, "bubbleChart") != null)
            {
                DrawScatterChart(g, chartDoc, rect, theme,
                    FindFirst(chartDoc, "bubbleChart") != null);
                return true;
            }
            if (FindFirst(chartDoc, "radarChart") != null)
            {
                DrawRadarChart(g, chartDoc, rect, theme);
                return true;
            }
            if (FindFirst(chartDoc, "areaChart") != null)
            {
                DrawAreaChart(g, chartDoc, rect, theme);
                return true;
            }

            DrawChart(g, chartDoc, rect, theme);
            return true;
        }

        private static Color[] EnhancedChartPalette(Dictionary<string, Color> theme)
        {
            return new Color[]
            {
                ThemeOrDefault(theme, "accent1", Color.FromArgb(79,129,189)),
                ThemeOrDefault(theme, "accent2", Color.FromArgb(192,80,77)),
                ThemeOrDefault(theme, "accent3", Color.FromArgb(155,187,89)),
                ThemeOrDefault(theme, "accent4", Color.FromArgb(128,100,162)),
                ThemeOrDefault(theme, "accent5", Color.FromArgb(75,172,198)),
                ThemeOrDefault(theme, "accent6", Color.FromArgb(247,150,70))
            };
        }

        private static List<ChartSeriesData> ReadStandardChartSeries(
            XmlDocument chartDoc,
            Dictionary<string, Color> theme)
        {
            List<ChartSeriesData> series =
                new List<ChartSeriesData>();

            List<XmlNode> nodes =
                FindAll(
                    chartDoc,
                    "ser");

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                ChartSeriesData item =
                    new ChartSeriesData();

                item.Name =
                    ReadChartSeriesName(
                        nodes[i]);

                ReadChartCategories(
                    nodes[i],
                    item.Categories);

                ReadChartValues(
                    nodes[i],
                    item.Values);

                item.ExplicitColor =
                    ReadChartSeriesColor(
                        nodes[i],
                        theme);

                ReadChartPointColors(
                    nodes[i],
                    theme,
                    item.PointColors,
                    item.Values.Count);

                ReadChartSeriesVisualStyle(
                    nodes[i],
                    item);

                if (item.Values.Count > 0)
                    series.Add(item);
            }

            return series;
        }

        private static void PrepareChartSurface(
            Graphics g,
            RectangleF rect,
            XmlDocument chartDoc,
            out RectangleF plot)
        {
            using (Brush bg =
                new SolidBrush(
                    Color.White))
            {
                g.FillRectangle(
                    bg,
                    rect);
            }

            using (Pen border =
                new Pen(
                    Color.FromArgb(
                        210,
                        210,
                        210),
                    1f))
            {
                g.DrawRectangle(
                    border,
                    rect.X,
                    rect.Y,
                    rect.Width,
                    rect.Height);
            }

            string title =
                ReadChartTitle(
                    chartDoc);

            float titleHeight =
                string.IsNullOrEmpty(
                    title)
                    ? 12f
                    : Math.Max(
                        34f,
                        rect.Height *
                        0.10f);

            if (!string.IsNullOrEmpty(
                    title))
            {
                using (Font font =
                    SafeFont(
                        "Arial",
                        Math.Max(
                            10f,
                            Math.Min(
                                18f,
                                rect.Height /
                                18f)),
                        FontStyle.Bold))
                using (Brush brush =
                    new SolidBrush(
                        Color.FromArgb(
                            45,
                            45,
                            45)))
                using (StringFormat sf =
                    new StringFormat())
                {
                    sf.Alignment =
                        StringAlignment.Center;
                    sf.LineAlignment =
                        StringAlignment.Center;

                    g.DrawString(
                        title,
                        font,
                        brush,
                        new RectangleF(
                            rect.Left + 5f,
                            rect.Top + 4f,
                            rect.Width - 10f,
                            titleHeight - 4f),
                        sf);
                }
            }

            string legendPosition =
                ReadChartLegendPosition(
                    chartDoc);

            float leftPad =
                Math.Max(
                    28f,
                    rect.Width *
                    0.07f);
            float rightPad =
                Math.Max(
                    28f,
                    rect.Width *
                    0.06f);
            float bottomPad =
                Math.Max(
                    30f,
                    rect.Height *
                    0.09f);
            float topPad =
                titleHeight;

            if (legendPosition == "r" ||
                legendPosition == "tr")
            {
                rightPad =
                    Math.Max(
                        82f,
                        rect.Width *
                        0.19f);
            }
            else if (legendPosition == "l")
            {
                leftPad =
                    Math.Max(
                        82f,
                        rect.Width *
                        0.19f);
            }
            else if (legendPosition == "b")
            {
                bottomPad =
                    Math.Max(
                        58f,
                        rect.Height *
                        0.17f);
            }
            else if (legendPosition == "t")
            {
                topPad +=
                    Math.Max(
                        28f,
                        rect.Height *
                        0.09f);
            }

            plot =
                new RectangleF(
                    rect.Left +
                        leftPad,
                    rect.Top +
                        topPad,
                    Math.Max(
                        12f,
                        rect.Width -
                            leftPad -
                            rightPad),
                    Math.Max(
                        12f,
                        rect.Height -
                            topPad -
                            bottomPad));
        }

        private static void DrawDoughnutChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            RectangleF plot;
            PrepareChartSurface(g, rect, chartDoc, out plot);
            List<ChartSeriesData> series = ReadStandardChartSeries(chartDoc, theme);
            if (series.Count == 0)
            {
                DrawPlaceholder(g, plot, "Doughnut chart");
                return;
            }

            ChartSeriesData data = series[0];
            double total = 0.0;
            for (int i = 0; i < data.Values.Count; i++)
                total += Math.Abs(data.Values[i]);
            if (total <= 0.0) total = 1.0;

            float size = Math.Min(plot.Width, plot.Height) * 0.88f;
            RectangleF pie = new RectangleF(
                plot.Left + (plot.Width - size) / 2f,
                plot.Top + (plot.Height - size) / 2f,
                size,
                size);
            Color[] palette = EnhancedChartPalette(theme);
            float angle =
                ReadChartFirstSliceAngle(
                    chartDoc);

            for (int i = 0; i < data.Values.Count; i++)
            {
                float sweep = (float)(360.0 * Math.Abs(data.Values[i]) / total);
                Color sliceColor =
                    GetChartPointColor(
                        data,
                        i,
                        palette);

                using (Brush brush =
                    new SolidBrush(
                        sliceColor))
                {
                    g.FillPie(
                        brush,
                        pie.X,
                        pie.Y,
                        pie.Width,
                        pie.Height,
                        angle,
                        sweep);
                }
                angle += sweep;
            }

            float holeSize =
                size *
                ReadDoughnutHoleRatio(
                    chartDoc);
            RectangleF hole = new RectangleF(
                pie.Left + (pie.Width - holeSize) / 2f,
                pie.Top + (pie.Height - holeSize) / 2f,
                holeSize,
                holeSize);
            using (Brush white =
                new SolidBrush(
                    Color.White))
            {
                g.FillEllipse(
                    white,
                    hole);
            }

            ChartLabelOptions labels =
                ReadChartLabelOptions(
                    chartDoc);

            if (labels.HasAny)
            {
                DrawDoughnutChartValueLabels(
                    g,
                    pie,
                    hole,
                    data,
                    labels,
                    ReadChartFirstSliceAngle(
                        chartDoc));
            }

            DrawChartLegend(
                g,
                rect,
                series,
                palette,
                "pie",
                ReadChartLegendPosition(chartDoc));
        }

        private static void DrawDoughnutChartValueLabels(
            Graphics g,
            RectangleF outer,
            RectangleF hole,
            ChartSeriesData series,
            ChartLabelOptions options,
            float startAngle)
        {
            if (series == null ||
                options == null ||
                series.Values.Count == 0)
            {
                return;
            }

            double total = 0.0;

            for (int i = 0;
                 i < series.Values.Count;
                 i++)
            {
                total +=
                    Math.Abs(
                        series.Values[i]);
            }

            if (total <= 0.0000001)
                return;

            float cx =
                outer.Left +
                outer.Width /
                2f;
            float cy =
                outer.Top +
                outer.Height /
                2f;

            float outerRadius =
                Math.Min(
                    outer.Width,
                    outer.Height) /
                2f;

            float innerRadius =
                Math.Min(
                    hole.Width,
                    hole.Height) /
                2f;

            float labelRadius =
                innerRadius +
                (outerRadius -
                 innerRadius) *
                0.55f;

            float angle =
                startAngle;

            using (Font font =
                SafeFont(
                    "Arial",
                    Math.Max(
                        7f,
                        Math.Min(
                            10f,
                            outer.Height /
                            28f)),
                    FontStyle.Bold))
            using (Brush brush =
                new SolidBrush(
                    Color.FromArgb(
                        55,
                        55,
                        55)))
            {
                for (int i = 0;
                     i < series.Values.Count;
                     i++)
                {
                    double value =
                        Math.Abs(
                            series.Values[i]);

                    float sweep =
                        (float)(
                            360.0 *
                            value /
                            total);

                    float centerAngle =
                        angle +
                        sweep /
                        2f;

                    double radians =
                        centerAngle *
                        Math.PI /
                        180.0;

                    string label =
                        BuildChartDataLabel(
                            options,
                            series,
                            i,
                            true);

                    if (!string.IsNullOrEmpty(
                            label))
                    {
                        SizeF size =
                            g.MeasureString(
                                label,
                                font);

                        float x =
                            cx +
                            (float)Math.Cos(
                                radians) *
                            labelRadius -
                            size.Width /
                            2f;

                        float y =
                            cy +
                            (float)Math.Sin(
                                radians) *
                            labelRadius -
                            size.Height /
                            2f;

                        g.DrawString(
                            label,
                            font,
                            brush,
                            x,
                            y);
                    }

                    angle +=
                        sweep;
                }
            }
        }

        private static void DrawAreaChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            RectangleF plot;
            PrepareChartSurface(
                g,
                rect,
                chartDoc,
                out plot);

            List<ChartSeriesData> series =
                ReadStandardChartSeries(
                    chartDoc,
                    theme);

            if (series.Count == 0)
            {
                DrawPlaceholder(
                    g,
                    plot,
                    "Area chart");
                return;
            }

            int categoryCount = 0;

            for (int s = 0;
                 s < series.Count;
                 s++)
            {
                categoryCount =
                    Math.Max(
                        categoryCount,
                        series[s].Values.Count);
            }

            if (categoryCount <= 0)
            {
                DrawPlaceholder(
                    g,
                    plot,
                    "Area chart");
                return;
            }

            string grouping =
                ReadAreaChartGrouping(
                    chartDoc);

            bool stacked =
                string.Equals(
                    grouping,
                    "stacked",
                    StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    grouping,
                    "percentStacked",
                    StringComparison.OrdinalIgnoreCase);

            bool percentStacked =
                string.Equals(
                    grouping,
                    "percentStacked",
                    StringComparison.OrdinalIgnoreCase);

            double minimum = 0.0;
            double maximum = 0.0;

            if (stacked)
            {
                CalculateStackedChartRange(
                    series,
                    categoryCount,
                    percentStacked,
                    out minimum,
                    out maximum);
            }
            else
            {
                for (int s = 0;
                     s < series.Count;
                     s++)
                {
                    for (int i = 0;
                         i < series[s].Values.Count;
                         i++)
                    {
                        double value =
                            series[s].Values[i];

                        minimum =
                            Math.Min(
                                minimum,
                                value);
                        maximum =
                            Math.Max(
                                maximum,
                                value);
                    }
                }

                if (Math.Abs(
                        maximum -
                        minimum) <
                    0.0000001)
                {
                    maximum =
                        minimum + 1.0;
                }
            }

            ChartAxisScale axisScale =
                ReadChartAxisScale(
                    chartDoc,
                    minimum,
                    maximum);

            if (percentStacked &&
                string.IsNullOrEmpty(
                    axisScale.NumberFormat))
            {
                axisScale.NumberFormat =
                    "0%";
            }

            DrawChartValueGrid(
                g,
                plot,
                axisScale,
                "column");

            float zeroY =
                plot.Bottom -
                (float)(
                    ChartAxisFraction(
                        0.0,
                        axisScale) *
                    plot.Height);

            zeroY =
                Math.Max(
                    plot.Top,
                    Math.Min(
                        plot.Bottom,
                        zeroY));

            using (Pen axis =
                new Pen(
                    Color.FromArgb(
                        100,
                        100,
                        100),
                    1.2f))
            {
                g.DrawLine(
                    axis,
                    plot.Left,
                    zeroY,
                    plot.Right,
                    zeroY);

                g.DrawLine(
                    axis,
                    plot.Left,
                    plot.Top,
                    plot.Left,
                    plot.Bottom);
            }

            Color[] palette =
                EnhancedChartPalette(
                    theme);

            ChartLabelOptions labels =
                ReadChartLabelOptions(
                    chartDoc);

            using (Font labelFont =
                SafeFont(
                    "Arial",
                    Math.Max(
                        6f,
                        Math.Min(
                            9f,
                            plot.Height /
                            36f))))
            using (Brush labelBrush =
                new SolidBrush(
                    Color.FromArgb(
                        70,
                        70,
                        70)))
            {
                for (int s = 0;
                     s < series.Count;
                     s++)
                {
                    ChartSeriesData item =
                        series[s];

                    if (item.Values.Count == 0)
                        continue;

                    List<PointF> topPoints =
                        new List<PointF>();
                    List<PointF> bottomPoints =
                        new List<PointF>();

                    for (int i = 0;
                         i < categoryCount;
                         i++)
                    {
                        double startValue = 0.0;
                        double endValue = 0.0;

                        if (i <
                            item.Values.Count)
                        {
                            if (stacked)
                            {
                                GetStackedChartSegment(
                                    series,
                                    i,
                                    s,
                                    percentStacked,
                                    out startValue,
                                    out endValue);
                            }
                            else
                            {
                                endValue =
                                    item.Values[i];
                            }
                        }

                        float x =
                            categoryCount <= 1
                                ? plot.Left +
                                  plot.Width /
                                  2f
                                : plot.Left +
                                  plot.Width *
                                  i /
                                  (categoryCount -
                                   1f);

                        float topY =
                            plot.Bottom -
                            (float)(
                                ChartAxisFraction(
                                    endValue,
                                    axisScale) *
                                plot.Height);

                        float bottomY =
                            plot.Bottom -
                            (float)(
                                ChartAxisFraction(
                                    startValue,
                                    axisScale) *
                                plot.Height);

                        topPoints.Add(
                            new PointF(
                                x,
                                topY));

                        bottomPoints.Add(
                            new PointF(
                                x,
                                bottomY));
                    }

                    List<PointF> polygon =
                        new List<PointF>();

                    for (int i = 0;
                         i < topPoints.Count;
                         i++)
                    {
                        polygon.Add(
                            topPoints[i]);
                    }

                    for (int i =
                             bottomPoints.Count -
                             1;
                         i >= 0;
                         i--)
                    {
                        polygon.Add(
                            bottomPoints[i]);
                    }

                    Color color =
                        GetChartSeriesColor(
                            series,
                            s,
                            palette);

                    if (polygon.Count >= 3)
                    {
                        using (Brush fill =
                            new SolidBrush(
                                Color.FromArgb(
                                    72,
                                    color)))
                        {
                            g.FillPolygon(
                                fill,
                                polygon.ToArray());
                        }
                    }

                    using (Pen pen =
                        new Pen(
                            color,
                            Math.Max(
                                1f,
                                item.LineWidth)))
                    {
                        pen.DashStyle =
                            item.LineDashStyle;

                        if (topPoints.Count > 1)
                        {
                            g.DrawLines(
                                pen,
                                topPoints.ToArray());
                        }
                    }

                    if (item.MarkerEnabled)
                    {
                        for (int i = 0;
                             i < topPoints.Count;
                             i++)
                        {
                            DrawChartMarker(
                                g,
                                topPoints[i],
                                item.MarkerSymbol,
                                item.MarkerSize,
                                color);
                        }
                    }

                    if (labels.HasAny)
                    {
                        for (int i = 0;
                             i < item.Values.Count &&
                             i < topPoints.Count;
                             i++)
                        {
                            string label =
                                BuildChartDataLabel(
                                    labels,
                                    item,
                                    i,
                                    false);

                            DrawChartValueLabel(
                                g,
                                labelFont,
                                labelBrush,
                                label,
                                topPoints[i].X,
                                topPoints[i].Y -
                                    labelFont.Height -
                                    2f,
                                true);
                        }
                    }
                }
            }

            DrawChartCategoryLabels(
                g,
                plot,
                series[0],
                categoryCount,
                "column");

            DrawChartAxisTitles(
                g,
                plot,
                rect,
                "column",
                ReadChartAxisTitle(
                    chartDoc,
                    "valAx"),
                ReadChartAxisTitle(
                    chartDoc,
                    "catAx"));

            DrawChartLegend(
                g,
                rect,
                series,
                palette,
                "area",
                ReadChartLegendPosition(
                    chartDoc));
        }

        private static string ReadAreaChartGrouping(
            XmlDocument chartDoc)
        {
            XmlNode areaChart =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "areaChart");

            if (areaChart == null)
                return "standard";

            XmlNode grouping =
                DirectChild(
                    areaChart,
                    "grouping");

            string value =
                grouping != null
                    ? GetAttr(
                        grouping,
                        "val")
                    : null;

            return string.IsNullOrEmpty(
                    value)
                ? "standard"
                : value;
        }

        private static void DrawScatterChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme,
            bool bubble)
        {
            RectangleF plot;
            PrepareChartSurface(
                g,
                rect,
                chartDoc,
                out plot);

            List<XmlNode> seriesNodes =
                FindAll(
                    chartDoc,
                    "ser");

            Color[] palette =
                EnhancedChartPalette(
                    theme);

            double minX = double.MaxValue;
            double maxX = double.MinValue;
            double minY = double.MaxValue;
            double maxY = double.MinValue;

            List<List<double>> allX =
                new List<List<double>>();
            List<List<double>> allY =
                new List<List<double>>();
            List<List<double>> allSize =
                new List<List<double>>();
            List<ChartSeriesData> styles =
                new List<ChartSeriesData>();

            for (int s = 0;
                 s < seriesNodes.Count;
                 s++)
            {
                List<double> xs =
                    ReadCachedNumbers(
                        DirectChild(
                            seriesNodes[s],
                            "xVal"));

                List<double> ys =
                    ReadCachedNumbers(
                        DirectChild(
                            seriesNodes[s],
                            "yVal"));

                List<double> sizes =
                    ReadCachedNumbers(
                        DirectChild(
                            seriesNodes[s],
                            "bubbleSize"));

                allX.Add(xs);
                allY.Add(ys);
                allSize.Add(sizes);

                ChartSeriesData style =
                    new ChartSeriesData();

                style.Name =
                    ReadChartSeriesName(
                        seriesNodes[s]);

                style.ExplicitColor =
                    ReadChartSeriesColor(
                        seriesNodes[s],
                        theme);

                ReadChartSeriesVisualStyle(
                    seriesNodes[s],
                    style);

                int pointCount =
                    Math.Min(
                        xs.Count,
                        ys.Count);

                ReadChartPointColors(
                    seriesNodes[s],
                    theme,
                    style.PointColors,
                    pointCount);

                styles.Add(style);

                for (int i = 0;
                     i < pointCount;
                     i++)
                {
                    minX =
                        Math.Min(
                            minX,
                            xs[i]);
                    maxX =
                        Math.Max(
                            maxX,
                            xs[i]);
                    minY =
                        Math.Min(
                            minY,
                            ys[i]);
                    maxY =
                        Math.Max(
                            maxY,
                            ys[i]);
                }
            }

            if (minX == double.MaxValue)
            {
                DrawPlaceholder(
                    g,
                    plot,
                    bubble
                        ? "Bubble chart"
                        : "Scatter chart");
                return;
            }

            if (Math.Abs(
                    maxX -
                    minX) < 0.000001)
            {
                maxX =
                    minX + 1.0;
            }

            if (Math.Abs(
                    maxY -
                    minY) < 0.000001)
            {
                maxY =
                    minY + 1.0;
            }

            using (Pen axis =
                new Pen(
                    Color.FromArgb(
                        110,
                        110,
                        110),
                    1f))
            {
                g.DrawRectangle(
                    axis,
                    plot.X,
                    plot.Y,
                    plot.Width,
                    plot.Height);
            }

            string scatterStyle =
                ReadScatterChartStyle(
                    chartDoc);

            bool drawSeriesLine =
                !bubble &&
                (scatterStyle.IndexOf(
                    "line",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                 scatterStyle.IndexOf(
                    "smooth",
                    StringComparison.OrdinalIgnoreCase) >= 0);

            bool smoothSeriesLine =
                !bubble &&
                scatterStyle.IndexOf(
                    "smooth",
                    StringComparison.OrdinalIgnoreCase) >= 0;

            bool drawMarkers =
                bubble ||
                scatterStyle.IndexOf(
                    "marker",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                string.IsNullOrEmpty(
                    scatterStyle);

            for (int s = 0;
                 s < allX.Count;
                 s++)
            {
                int pointCount =
                    Math.Min(
                        allX[s].Count,
                        allY[s].Count);

                if (pointCount <= 0)
                    continue;

                ChartSeriesData style =
                    s < styles.Count
                        ? styles[s]
                        : new ChartSeriesData();

                Color seriesColor =
                    style.ExplicitColor ??
                    palette[
                        s %
                        palette.Length];

                PointF[] points =
                    new PointF[
                        pointCount];

                for (int i = 0;
                     i < pointCount;
                     i++)
                {
                    float x =
                        plot.Left +
                        (float)(
                            (allX[s][i] -
                             minX) /
                            (maxX -
                             minX)) *
                        plot.Width;

                    float y =
                        plot.Bottom -
                        (float)(
                            (allY[s][i] -
                             minY) /
                            (maxY -
                             minY)) *
                        plot.Height;

                    points[i] =
                        new PointF(
                            x,
                            y);
                }

                if (drawSeriesLine &&
                    points.Length > 1)
                {
                    using (Pen pen =
                        new Pen(
                            seriesColor,
                            Math.Max(
                                1f,
                                style.LineWidth)))
                    {
                        pen.DashStyle =
                            style.LineDashStyle;

                        if (smoothSeriesLine &&
                            points.Length > 2)
                        {
                            g.DrawCurve(
                                pen,
                                points,
                                0.5f);
                        }
                        else
                        {
                            g.DrawLines(
                                pen,
                                points);
                        }
                    }
                }

                for (int i = 0;
                     i < pointCount;
                     i++)
                {
                    Color pointColor =
                        GetChartPointColor(
                            style,
                            i,
                            palette);

                    if (bubble)
                    {
                        float radius = 4f;

                        if (i <
                            allSize[s].Count)
                        {
                            radius =
                                Math.Max(
                                    3f,
                                    Math.Min(
                                        20f,
                                        (float)Math.Sqrt(
                                            Math.Abs(
                                                allSize[s][i]))));
                        }

                        using (Brush brush =
                            new SolidBrush(
                                Color.FromArgb(
                                    150,
                                    pointColor)))
                        using (Pen pen =
                            new Pen(
                                pointColor,
                                1f))
                        {
                            g.FillEllipse(
                                brush,
                                points[i].X -
                                    radius,
                                points[i].Y -
                                    radius,
                                radius * 2f,
                                radius * 2f);

                            g.DrawEllipse(
                                pen,
                                points[i].X -
                                    radius,
                                points[i].Y -
                                    radius,
                                radius * 2f,
                                radius * 2f);
                        }
                    }
                    else if (drawMarkers &&
                             style.MarkerEnabled)
                    {
                        DrawChartMarker(
                            g,
                            points[i],
                            style.MarkerSymbol,
                            style.MarkerSize,
                            pointColor);
                    }
                }
            }

            if (!bubble)
            {
                List<ChartSeriesData> legendSeries =
                    styles;

                DrawChartLegend(
                    g,
                    rect,
                    legendSeries,
                    palette,
                    "scatter",
                    ReadChartLegendPosition(
                        chartDoc));
            }
        }

        private static string ReadScatterChartStyle(
            XmlDocument chartDoc)
        {
            XmlNode scatter =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "scatterChart");

            if (scatter == null)
                return string.Empty;

            XmlNode style =
                DirectChild(
                    scatter,
                    "scatterStyle");

            return style != null
                ? GetAttr(
                    style,
                    "val") ??
                  string.Empty
                : string.Empty;
        }

        private static void DrawRadarChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            RectangleF plot;
            PrepareChartSurface(g, rect, chartDoc, out plot);
            List<ChartSeriesData> series = ReadStandardChartSeries(chartDoc, theme);
            if (series.Count == 0)
            {
                DrawPlaceholder(g, plot, "Radar chart");
                return;
            }

            int count = 0;
            double max = 0.0;
            for (int s = 0; s < series.Count; s++)
            {
                count = Math.Max(count, series[s].Values.Count);
                for (int i = 0; i < series[s].Values.Count; i++)
                    max = Math.Max(max, Math.Abs(series[s].Values[i]));
            }
            if (count < 3 || max <= 0.0)
            {
                DrawChart(g, chartDoc, rect, theme);
                return;
            }

            float cx = plot.Left + plot.Width / 2f;
            float cy = plot.Top + plot.Height / 2f;
            float radius = Math.Min(plot.Width, plot.Height) * 0.43f;
            using (Pen grid = new Pen(Color.FromArgb(190, 195, 200), 1f))
            {
                for (int ring = 1; ring <= 4; ring++)
                {
                    PointF[] ringPoints = new PointF[count];
                    for (int i = 0; i < count; i++)
                    {
                        double angle = -Math.PI / 2.0 + i * Math.PI * 2.0 / count;
                        float rr = radius * ring / 4f;
                        ringPoints[i] = new PointF(
                            cx + (float)Math.Cos(angle) * rr,
                            cy + (float)Math.Sin(angle) * rr);
                    }
                    g.DrawPolygon(grid, ringPoints);
                }
                for (int i = 0; i < count; i++)
                {
                    double angle = -Math.PI / 2.0 + i * Math.PI * 2.0 / count;
                    g.DrawLine(grid, cx, cy,
                        cx + (float)Math.Cos(angle) * radius,
                        cy + (float)Math.Sin(angle) * radius);
                }
            }

            Color[] palette = EnhancedChartPalette(theme);
            for (int s = 0; s < series.Count; s++)
            {
                if (series[s].Values.Count == 0) continue;
                PointF[] points = new PointF[count];
                for (int i = 0; i < count; i++)
                {
                    double value = i < series[s].Values.Count ? Math.Abs(series[s].Values[i]) : 0.0;
                    float rr = radius * (float)(value / max);
                    double angle = -Math.PI / 2.0 + i * Math.PI * 2.0 / count;
                    points[i] = new PointF(
                        cx + (float)Math.Cos(angle) * rr,
                        cy + (float)Math.Sin(angle) * rr);
                }
                Color color =
                    GetChartSeriesColor(
                        series,
                        s,
                        palette);

                using (Brush brush =
                    new SolidBrush(
                        Color.FromArgb(
                            42,
                            color)))
                {
                    g.FillPolygon(
                        brush,
                        points);
                }

                using (Pen pen =
                    new Pen(
                        color,
                        Math.Max(
                            1f,
                            series[s].LineWidth)))
                {
                    pen.DashStyle =
                        series[s].LineDashStyle;
                    g.DrawPolygon(
                        pen,
                        points);
                }

                if (series[s].MarkerEnabled)
                {
                    for (int i = 0;
                         i < points.Length;
                         i++)
                    {
                        DrawChartMarker(
                            g,
                            points[i],
                            series[s].MarkerSymbol,
                            series[s].MarkerSize,
                            color);
                    }
                }
            }
            DrawRadarCategoryLabels(
                g,
                series[0],
                count,
                cx,
                cy,
                radius);

            DrawChartLegend(
                g,
                rect,
                series,
                palette,
                "radar",
                ReadChartLegendPosition(chartDoc));
        }

        private static void DrawRadarCategoryLabels(
            Graphics g,
            ChartSeriesData series,
            int count,
            float cx,
            float cy,
            float radius)
        {
            if (series == null ||
                series.Categories.Count == 0 ||
                count <= 0)
            {
                return;
            }

            using (Font font =
                SafeFont(
                    "Arial",
                    8f))
            using (Brush brush =
                new SolidBrush(
                    Color.FromArgb(
                        72,
                        72,
                        72)))
            {
                int labelCount =
                    Math.Min(
                        count,
                        series.Categories.Count);

                for (int i = 0;
                     i < labelCount;
                     i++)
                {
                    string label =
                        series.Categories[i];

                    if (string.IsNullOrEmpty(
                            label))
                    {
                        continue;
                    }

                    double angle =
                        -Math.PI /
                        2.0 +
                        i *
                        Math.PI *
                        2.0 /
                        count;

                    SizeF size =
                        g.MeasureString(
                            label,
                            font);

                    float x =
                        cx +
                        (float)Math.Cos(
                            angle) *
                        (radius + 10f);

                    float y =
                        cy +
                        (float)Math.Sin(
                            angle) *
                        (radius + 10f);

                    x -=
                        size.Width /
                        2f;
                    y -=
                        size.Height /
                        2f;

                    g.DrawString(
                        label,
                        font,
                        brush,
                        x,
                        y);
                }
            }
        }

        private static List<double> ReadCachedNumbers(XmlNode parent)
        {
            List<double> result = new List<double>();
            XmlNode cache = FindFirst(parent, "numCache");
            if (cache == null)
                cache = FindFirst(parent, "numLit");
            if (cache == null)
                return result;

            List<XmlNode> points = FindAll(cache, "pt");
            for (int i = 0; i < points.Count; i++)
            {
                XmlNode value = DirectChild(points[i], "v");
                double parsed;
                if (value != null && double.TryParse(
                    value.InnerText,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out parsed))
                {
                    result.Add(parsed);
                }
            }
            return result;
        }

        private static bool DrawStructuredSmartArt(
            Graphics g,
            XmlDocument dataDoc,
            XmlDocument layoutDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            List<XmlNode> points =
                FindAll(dataDoc, "pt");
            List<XmlNode> connections =
                FindAll(dataDoc, "cxn");

            Dictionary<string, SmartNode> nodes =
                new Dictionary<string, SmartNode>(
                    StringComparer.Ordinal);
            List<SmartNode> orderedNodes =
                new List<SmartNode>();
            HashSet<string> incoming =
                new HashSet<string>(
                    StringComparer.Ordinal);

            for (int i = 0; i < points.Count; i++)
            {
                string id =
                    GetAttr(
                        points[i],
                        "modelId");

                if (string.IsNullOrEmpty(id))
                    continue;

                string type =
                    GetAttr(
                        points[i],
                        "type");

                if (type == "doc" ||
                    type == "asst")
                {
                    continue;
                }

                SmartNode node =
                    new SmartNode();
                node.Id = id;
                node.Label =
                    ReadSmartArtLabel(
                        points[i]);

                if (string.IsNullOrEmpty(
                        node.Label))
                {
                    node.Label = "•";
                }

                nodes[id] = node;
                orderedNodes.Add(node);
            }

            for (int i = 0;
                 i < connections.Count;
                 i++)
            {
                string source =
                    GetAttr(
                        connections[i],
                        "srcId");
                string destination =
                    GetAttr(
                        connections[i],
                        "destId");

                if (string.IsNullOrEmpty(source) ||
                    string.IsNullOrEmpty(destination) ||
                    !nodes.ContainsKey(source) ||
                    !nodes.ContainsKey(destination) ||
                    source == destination)
                {
                    continue;
                }

                if (!nodes[source].Children.Contains(
                        destination))
                {
                    nodes[source].Children.Add(
                        destination);
                }

                incoming.Add(
                    destination);
            }

            if (nodes.Count < 2)
                return false;

            List<SmartNode> roots =
                new List<SmartNode>();

            for (int i = 0;
                 i < orderedNodes.Count;
                 i++)
            {
                if (!incoming.Contains(
                        orderedNodes[i].Id))
                {
                    roots.Add(
                        orderedNodes[i]);
                }
            }

            if (roots.Count == 0 &&
                orderedNodes.Count > 0)
            {
                roots.Add(
                    orderedNodes[0]);
            }

            Queue<SmartNode> queue =
                new Queue<SmartNode>();
            HashSet<string> visited =
                new HashSet<string>(
                    StringComparer.Ordinal);

            for (int i = 0;
                 i < roots.Count;
                 i++)
            {
                roots[i].Depth = 0;
                queue.Enqueue(
                    roots[i]);
            }

            while (queue.Count > 0)
            {
                SmartNode current =
                    queue.Dequeue();

                if (!visited.Add(
                        current.Id))
                {
                    continue;
                }

                for (int i = 0;
                     i < current.Children.Count;
                     i++)
                {
                    SmartNode child;
                    if (!nodes.TryGetValue(
                            current.Children[i],
                            out child))
                    {
                        continue;
                    }

                    child.Depth =
                        Math.Max(
                            child.Depth,
                            current.Depth + 1);

                    queue.Enqueue(
                        child);
                }
            }

            int maxDepth = 0;
            for (int i = 0;
                 i < orderedNodes.Count;
                 i++)
            {
                maxDepth =
                    Math.Max(
                        maxDepth,
                        orderedNodes[i].Depth);
            }

            string layoutKind =
                ReadSmartArtLayoutKind(
                    layoutDoc);

            Dictionary<string, RectangleF> positions =
                BuildSmartArtPositions(
                    layoutKind,
                    orderedNodes,
                    rect,
                    maxDepth);

            if (positions.Count == 0)
                return false;

            Color[] palette =
                EnhancedChartPalette(
                    theme);

            DrawSmartArtConnectors(
                g,
                layoutKind,
                orderedNodes,
                nodes,
                positions,
                palette);

            for (int i = 0;
                 i < orderedNodes.Count;
                 i++)
            {
                SmartNode node =
                    orderedNodes[i];

                RectangleF box;
                if (!positions.TryGetValue(
                        node.Id,
                        out box))
                {
                    continue;
                }

                Color accent =
                    palette[
                        Math.Abs(
                            node.Depth) %
                        palette.Length];

                Color fill =
                    LightenSmartArtColor(
                        accent,
                        0.82f);

                if (layoutKind == "venn")
                {
                    fill =
                        Color.FromArgb(
                            145,
                            fill);
                }

                bool ellipseNode =
                    layoutKind == "cycle" ||
                    layoutKind == "radial" ||
                    layoutKind == "venn";

                using (GraphicsPath path =
                    ellipseNode
                        ? EllipsePath(box)
                        : RoundedRectanglePath(
                            box,
                            Math.Min(
                                12f,
                                box.Height *
                                0.24f)))
                using (Brush fillBrush =
                    new SolidBrush(fill))
                using (Pen border =
                    new Pen(
                        accent,
                        1.6f))
                {
                    g.FillPath(
                        fillBrush,
                        path);
                    g.DrawPath(
                        border,
                        path);
                }

                using (Font font = SafeFont(
                    "Arial",
                    Math.Max(
                        7f,
                        Math.Min(
                            13f,
                            box.Height /
                            4.2f)),
                    node.Depth == 0
                        ? FontStyle.Bold
                        : FontStyle.Regular))
                using (Brush text =
                    new SolidBrush(
                        Color.FromArgb(
                            42,
                            47,
                            56)))
                using (StringFormat sf =
                    new StringFormat())
                {
                    sf.Alignment =
                        StringAlignment.Center;
                    sf.LineAlignment =
                        StringAlignment.Center;
                    sf.Trimming =
                        StringTrimming.EllipsisCharacter;

                    g.DrawString(
                        node.Label,
                        font,
                        text,
                        box,
                        sf);
                }
            }

            return true;
        }

        private static string ReadSmartArtLayoutKind(
            XmlDocument layoutDoc)
        {
            if (layoutDoc == null ||
                layoutDoc.DocumentElement == null)
            {
                return "hierarchy";
            }

            string text =
                layoutDoc.OuterXml
                    .ToLowerInvariant();

            if (text.IndexOf("venn") >= 0)
                return "venn";

            if (text.IndexOf("radial") >= 0 ||
                text.IndexOf("relationship") >= 0 ||
                text.IndexOf("converging") >= 0 ||
                text.IndexOf("diverging") >= 0)
            {
                return "radial";
            }

            if ((text.IndexOf("vertical") >= 0 ||
                 text.IndexOf("descending") >= 0) &&
                (text.IndexOf("process") >= 0 ||
                 text.IndexOf("step") >= 0))
            {
                return "verticalProcess";
            }

            if (text.IndexOf("cycle") >= 0 ||
                text.IndexOf("circular") >= 0 ||
                text.IndexOf("continuous loop") >= 0)
            {
                return "cycle";
            }

            if (text.IndexOf("pyramid") >= 0 ||
                text.IndexOf("funnel") >= 0)
            {
                return "pyramid";
            }

            if (text.IndexOf("matrix") >= 0 ||
                text.IndexOf("grid") >= 0 ||
                text.IndexOf("quadrant") >= 0)
            {
                return "matrix";
            }

            if (text.IndexOf("list") >= 0 ||
                text.IndexOf("bullet") >= 0 ||
                text.IndexOf("stacked") >= 0)
            {
                return "list";
            }

            if (text.IndexOf("process") >= 0 ||
                text.IndexOf("chevron") >= 0 ||
                text.IndexOf("step") >= 0 ||
                text.IndexOf("timeline") >= 0)
            {
                return "process";
            }

            if (text.IndexOf("hierarchy") >= 0 ||
                text.IndexOf("orgchart") >= 0 ||
                text.IndexOf("organization") >= 0 ||
                text.IndexOf("horizontal hierarchy") >= 0)
            {
                return "hierarchy";
            }

            return "hierarchy";
        }

        private static Dictionary<string, RectangleF>
            BuildSmartArtPositions(
                string layoutKind,
                List<SmartNode> nodes,
                RectangleF rect,
                int maxDepth)
        {
            if (layoutKind == "cycle")
            {
                return BuildCycleSmartArtPositions(
                    nodes,
                    rect);
            }

            if (layoutKind == "process")
            {
                return BuildProcessSmartArtPositions(
                    nodes,
                    rect);
            }

            if (layoutKind == "verticalProcess")
            {
                return BuildVerticalProcessSmartArtPositions(
                    nodes,
                    rect);
            }

            if (layoutKind == "list")
            {
                return BuildListSmartArtPositions(
                    nodes,
                    rect);
            }

            if (layoutKind == "radial")
            {
                return BuildRadialSmartArtPositions(
                    nodes,
                    rect);
            }

            if (layoutKind == "venn")
            {
                return BuildVennSmartArtPositions(
                    nodes,
                    rect);
            }

            if (layoutKind == "matrix")
            {
                return BuildMatrixSmartArtPositions(
                    nodes,
                    rect);
            }

            if (layoutKind == "pyramid")
            {
                return BuildPyramidSmartArtPositions(
                    nodes,
                    rect);
            }

            return BuildHierarchySmartArtPositions(
                nodes,
                rect,
                maxDepth);
        }

        private static Dictionary<string, RectangleF>
            BuildHierarchySmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect,
                int maxDepth)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            int rows =
                Math.Max(
                    1,
                    maxDepth + 1);

            float rowHeight =
                rect.Height /
                rows;

            for (int depth = 0;
                 depth < rows;
                 depth++)
            {
                List<SmartNode> level =
                    new List<SmartNode>();

                for (int i = 0;
                     i < nodes.Count;
                     i++)
                {
                    if (nodes[i].Depth == depth)
                        level.Add(nodes[i]);
                }

                if (level.Count == 0)
                    continue;

                float columnWidth =
                    rect.Width /
                    level.Count;

                float boxW =
                    Math.Max(
                        52f,
                        columnWidth *
                        0.72f);

                float boxH =
                    Math.Max(
                        30f,
                        rowHeight *
                        0.52f);

                for (int i = 0;
                     i < level.Count;
                     i++)
                {
                    float x =
                        rect.Left +
                        i * columnWidth +
                        (columnWidth -
                         boxW) /
                        2f;

                    float y =
                        rect.Top +
                        depth * rowHeight +
                        (rowHeight -
                         boxH) /
                        2f;

                    positions[
                        level[i].Id] =
                        new RectangleF(
                            x,
                            y,
                            boxW,
                            boxH);
                }
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildProcessSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            int count =
                Math.Max(
                    1,
                    nodes.Count);

            int columns =
                Math.Min(
                    5,
                    count);

            int rows =
                (int)Math.Ceiling(
                    count /
                    (double)columns);

            float gap =
                Math.Max(
                    8f,
                    Math.Min(
                        rect.Width,
                        rect.Height) *
                    0.025f);

            float cellW =
                Math.Max(
                    40f,
                    (rect.Width -
                     gap * (columns + 1)) /
                    columns);

            float cellH =
                Math.Max(
                    28f,
                    (rect.Height -
                     gap * (rows + 1)) /
                    rows);

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                int row =
                    i /
                    columns;
                int column =
                    i %
                    columns;

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        rect.Left +
                            gap +
                            column *
                            (cellW + gap),
                        rect.Top +
                            gap +
                            row *
                            (cellH + gap),
                        cellW,
                        cellH);
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildVerticalProcessSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            int count =
                Math.Max(
                    1,
                    nodes.Count);

            float gap =
                Math.Max(
                    7f,
                    rect.Height *
                    0.025f);

            float boxH =
                Math.Max(
                    28f,
                    (rect.Height -
                     gap *
                     (count + 1)) /
                    count);

            boxH =
                Math.Min(
                    boxH,
                    rect.Height *
                    0.22f);

            float boxW =
                Math.Max(
                    60f,
                    rect.Width *
                    0.72f);

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                float x =
                    rect.Left +
                    (rect.Width -
                     boxW) /
                    2f;

                float y =
                    rect.Top +
                    gap +
                    i *
                    (boxH + gap);

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        x,
                        y,
                        boxW,
                        boxH);
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildListSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            int count =
                Math.Max(
                    1,
                    nodes.Count);

            int columns =
                count > 6
                    ? 2
                    : 1;

            int rows =
                (int)Math.Ceiling(
                    count /
                    (double)columns);

            float gap =
                Math.Max(
                    8f,
                    Math.Min(
                        rect.Width,
                        rect.Height) *
                    0.025f);

            float boxW =
                Math.Max(
                    70f,
                    (rect.Width -
                     gap *
                     (columns + 1)) /
                    columns);

            float boxH =
                Math.Max(
                    28f,
                    (rect.Height -
                     gap *
                     (rows + 1)) /
                    rows);

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                int column =
                    i /
                    rows;
                int row =
                    i %
                    rows;

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        rect.Left +
                            gap +
                            column *
                            (boxW + gap),
                        rect.Top +
                            gap +
                            row *
                            (boxH + gap),
                        boxW,
                        boxH);
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildRadialSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            if (nodes.Count == 0)
                return positions;

            float centerW =
                Math.Max(
                    64f,
                    Math.Min(
                        rect.Width *
                        0.30f,
                        180f));

            float centerH =
                Math.Max(
                    38f,
                    Math.Min(
                        rect.Height *
                        0.20f,
                        100f));

            float cx =
                rect.Left +
                rect.Width /
                2f;
            float cy =
                rect.Top +
                rect.Height /
                2f;

            positions[
                nodes[0].Id] =
                new RectangleF(
                    cx -
                        centerW /
                        2f,
                    cy -
                        centerH /
                        2f,
                    centerW,
                    centerH);

            if (nodes.Count == 1)
                return positions;

            float satelliteW =
                Math.Max(
                    52f,
                    centerW *
                    0.78f);
            float satelliteH =
                Math.Max(
                    32f,
                    centerH *
                    0.84f);

            float radiusX =
                Math.Max(
                    12f,
                    (rect.Width -
                     satelliteW) *
                    0.40f);
            float radiusY =
                Math.Max(
                    12f,
                    (rect.Height -
                     satelliteH) *
                    0.38f);

            int satelliteCount =
                nodes.Count -
                1;

            for (int i = 1;
                 i < nodes.Count;
                 i++)
            {
                double angle =
                    -Math.PI /
                    2.0 +
                    (i - 1) *
                    Math.PI *
                    2.0 /
                    satelliteCount;

                float x =
                    cx +
                    (float)Math.Cos(
                        angle) *
                    radiusX -
                    satelliteW /
                    2f;

                float y =
                    cy +
                    (float)Math.Sin(
                        angle) *
                    radiusY -
                    satelliteH /
                    2f;

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        x,
                        y,
                        satelliteW,
                        satelliteH);
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildVennSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            if (nodes.Count == 0)
                return positions;

            int count =
                Math.Min(
                    6,
                    nodes.Count);

            float diameter =
                Math.Max(
                    64f,
                    Math.Min(
                        rect.Width,
                        rect.Height) *
                    (count <= 3
                        ? 0.42f
                        : 0.32f));

            float cx =
                rect.Left +
                rect.Width /
                2f;
            float cy =
                rect.Top +
                rect.Height /
                2f;

            float ring =
                diameter *
                (count <= 3
                    ? 0.34f
                    : 0.58f);

            for (int i = 0;
                 i < count;
                 i++)
            {
                double angle =
                    -Math.PI /
                    2.0 +
                    i *
                    Math.PI *
                    2.0 /
                    count;

                float x =
                    cx +
                    (float)Math.Cos(
                        angle) *
                    ring -
                    diameter /
                    2f;

                float y =
                    cy +
                    (float)Math.Sin(
                        angle) *
                    ring -
                    diameter /
                    2f;

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        x,
                        y,
                        diameter,
                        diameter);
            }

            for (int i = count;
                 i < nodes.Count;
                 i++)
            {
                float small =
                    Math.Max(
                        42f,
                        diameter *
                        0.62f);

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        cx -
                            small /
                            2f,
                        cy -
                            small /
                            2f,
                        small,
                        small);
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildCycleSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            if (nodes.Count == 0)
                return positions;

            float boxW =
                Math.Max(
                    54f,
                    Math.Min(
                        rect.Width * 0.23f,
                        180f));

            float boxH =
                Math.Max(
                    32f,
                    Math.Min(
                        rect.Height * 0.17f,
                        96f));

            float cx =
                rect.Left +
                rect.Width /
                2f;
            float cy =
                rect.Top +
                rect.Height /
                2f;

            float rx =
                Math.Max(
                    8f,
                    (rect.Width -
                     boxW) *
                    0.40f);

            float ry =
                Math.Max(
                    8f,
                    (rect.Height -
                     boxH) *
                    0.38f);

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                double angle =
                    -Math.PI /
                    2.0 +
                    i *
                    Math.PI *
                    2.0 /
                    nodes.Count;

                float x =
                    cx +
                    (float)Math.Cos(
                        angle) *
                    rx -
                    boxW /
                    2f;

                float y =
                    cy +
                    (float)Math.Sin(
                        angle) *
                    ry -
                    boxH /
                    2f;

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        x,
                        y,
                        boxW,
                        boxH);
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildMatrixSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            int count =
                Math.Max(
                    1,
                    nodes.Count);

            int columns =
                (int)Math.Ceiling(
                    Math.Sqrt(count));

            int rows =
                (int)Math.Ceiling(
                    count /
                    (double)columns);

            float gap =
                Math.Max(
                    8f,
                    Math.Min(
                        rect.Width,
                        rect.Height) *
                    0.03f);

            float cellW =
                Math.Max(
                    36f,
                    (rect.Width -
                     gap *
                     (columns + 1)) /
                    columns);

            float cellH =
                Math.Max(
                    28f,
                    (rect.Height -
                     gap *
                     (rows + 1)) /
                    rows);

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                int row =
                    i /
                    columns;
                int column =
                    i %
                    columns;

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        rect.Left +
                            gap +
                            column *
                            (cellW + gap),
                        rect.Top +
                            gap +
                            row *
                            (cellH + gap),
                        cellW,
                        cellH);
            }

            return positions;
        }

        private static Dictionary<string, RectangleF>
            BuildPyramidSmartArtPositions(
                List<SmartNode> nodes,
                RectangleF rect)
        {
            Dictionary<string, RectangleF> positions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);

            int count =
                Math.Max(
                    1,
                    nodes.Count);

            float rowHeight =
                rect.Height /
                count;

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                float ratio =
                    count <= 1
                        ? 0f
                        : i /
                            (float)(
                                count - 1);

                float width =
                    rect.Width *
                    (0.40f +
                     ratio *
                     0.48f);

                float height =
                    Math.Max(
                        26f,
                        rowHeight *
                        0.70f);

                float x =
                    rect.Left +
                    (rect.Width -
                     width) /
                    2f;

                float y =
                    rect.Top +
                    i *
                    rowHeight +
                    (rowHeight -
                     height) /
                    2f;

                positions[
                    nodes[i].Id] =
                    new RectangleF(
                        x,
                        y,
                        width,
                        height);
            }

            return positions;
        }

        private static void DrawSmartArtConnectors(
            Graphics g,
            string layoutKind,
            List<SmartNode> orderedNodes,
            Dictionary<string, SmartNode> nodes,
            Dictionary<string, RectangleF> positions,
            Color[] palette)
        {
            if (layoutKind == "venn")
                return;

            Color connectorColor =
                palette != null &&
                palette.Length > 0
                    ? Color.FromArgb(
                        135,
                        palette[0])
                    : Color.FromArgb(
                        120,
                        79,
                        129,
                        189);

            using (Pen connector =
                new Pen(
                    connectorColor,
                    1.6f))
            {
                connector.EndCap =
                    LineCap.ArrowAnchor;

                for (int n = 0;
                     n < orderedNodes.Count;
                     n++)
                {
                    SmartNode node =
                        orderedNodes[n];

                    RectangleF source;
                    if (!positions.TryGetValue(
                            node.Id,
                            out source))
                    {
                        continue;
                    }

                    for (int i = 0;
                         i < node.Children.Count;
                         i++)
                    {
                        RectangleF destination;
                        if (!positions.TryGetValue(
                                node.Children[i],
                                out destination))
                        {
                            continue;
                        }

                        PointF from;
                        PointF to;

                        if (layoutKind == "hierarchy" ||
                            layoutKind == "verticalProcess")
                        {
                            from =
                                new PointF(
                                    source.Left +
                                    source.Width /
                                    2f,
                                    source.Bottom);

                            to =
                                new PointF(
                                    destination.Left +
                                    destination.Width /
                                    2f,
                                    destination.Top);
                        }
                        else
                        {
                            from =
                                new PointF(
                                    source.Left +
                                    source.Width /
                                    2f,
                                    source.Top +
                                    source.Height /
                                    2f);

                            to =
                                new PointF(
                                    destination.Left +
                                    destination.Width /
                                    2f,
                                    destination.Top +
                                    destination.Height /
                                    2f);
                        }

                        g.DrawLine(
                            connector,
                            from,
                            to);
                    }
                }
            }
        }

        private static Color LightenSmartArtColor(
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
                255,
                color.R +
                    (int)(
                        (255 -
                         color.R) *
                        amount),
                color.G +
                    (int)(
                        (255 -
                         color.G) *
                        amount),
                color.B +
                    (int)(
                        (255 -
                         color.B) *
                        amount));
        }

        private static GraphicsPath EllipsePath(
            RectangleF rect)
        {
            GraphicsPath path =
                new GraphicsPath();

            path.AddEllipse(rect);
            path.CloseFigure();
            return path;
        }

        private static string ReadSmartArtLabel(XmlNode point)
        {
            if (point == null)
                return string.Empty;
            List<XmlNode> texts = FindAll(point, "t");
            for (int i = 0; i < texts.Count; i++)
            {
                string value = texts[i].InnerText;
                if (!string.IsNullOrEmpty(value))
                    return value.Trim();
            }
            return string.Empty;
        }

        private static GraphicsPath RoundedRectanglePath(RectangleF rect, float radius)
        {
            GraphicsPath path = new GraphicsPath();
            float d = Math.Max(1f, radius * 2f);
            path.AddArc(rect.Left, rect.Top, d, d, 180f, 90f);
            path.AddArc(rect.Right - d, rect.Top, d, d, 270f, 90f);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0f, 90f);
            path.AddArc(rect.Left, rect.Bottom - d, d, d, 90f, 90f);
            path.CloseFigure();
            return path;
        }

        private static void DrawEnhancedGroup(
            ZipArchive zip,
            Graphics g,
            XmlNode group,
            Dictionary<string, string> rels,
            TransformContext ctx,
            Dictionary<string, Color> theme,
            bool inheritedLayer,
            Dictionary<string, RectangleF> placeholderRects,
            HashSet<string> hiddenShapeIds,
            int slideNumber)
        {
            TransformContext childContext = BuildGroupContext(group, ctx);
            RectangleF groupRect;
            bool hasRect = TryGetGroupRect(group, ctx, out groupRect);
            GraphicsState state = g.Save();

            try
            {
                if (hasRect)
                    ApplyRotation(g, group, groupRect);

                DrawContainer(
                    zip,
                    g,
                    group,
                    rels,
                    childContext,
                    theme,
                    inheritedLayer,
                    placeholderRects,
                    hiddenShapeIds,
                    slideNumber);
            }
            finally
            {
                g.Restore(state);
            }
        }

        private static bool TryGetGroupRect(
            XmlNode group,
            TransformContext parent,
            out RectangleF rect)
        {
            rect = RectangleF.Empty;
            XmlNode grpSpPr = DirectChild(group, "grpSpPr");
            XmlNode xfrm = grpSpPr == null ? null : DirectChild(grpSpPr, "xfrm");
            XmlNode off = xfrm == null ? null : DirectChild(xfrm, "off");
            XmlNode ext = xfrm == null ? null : DirectChild(xfrm, "ext");
            if (off == null || ext == null)
                return false;

            long x = GetLong(off, "x", 0);
            long y = GetLong(off, "y", 0);
            long w = Math.Max(1, GetLong(ext, "cx", 1));
            long h = Math.Max(1, GetLong(ext, "cy", 1));
            rect = new RectangleF(
                (float)(parent.Ax * x + parent.Bx),
                (float)(parent.Ay * y + parent.By),
                Math.Abs((float)(parent.Ax * w)),
                Math.Abs((float)(parent.Ay * h)));
            return true;
        }
    }
}
