using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
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
            public bool IsAssistant;
            public string ParentId;
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
                    name == "filter" ||
                    name == "symbol")
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

                if (name == "use" &&
                    TryDrawEnhancedSvgUseContainer(
                        g,
                        node,
                        target,
                        minX,
                        minY,
                        sx,
                        sy))
                {
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

                        if (DrawSvgColorMatrixApproximation(
                                g,
                                node,
                                parent.OwnerDocument,
                                path,
                                target,
                                sx,
                                sy))
                        {
                            drew = true;
                            continue;
                        }

                        if (DrawSvgOffsetGaussianChainApproximation(
                                g,
                                node,
                                parent.OwnerDocument,
                                path,
                                sx,
                                sy))
                        {
                            drew = true;
                            continue;
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

        private static bool TryDrawEnhancedSvgUseContainer(
            Graphics g,
            XmlNode useNode,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            if (g == null ||
                useNode == null ||
                useNode.LocalName != "use" ||
                HasSvgUseExpansionMarker(
                    useNode))
            {
                return false;
            }

            string referenceId =
                ReadSvgUseReferenceId(
                    useNode);

            if (string.IsNullOrEmpty(
                    referenceId))
            {
                return false;
            }

            XmlNode referenced =
                FindSvgNodeById(
                    useNode.OwnerDocument,
                    referenceId);

            if (referenced == null ||
                referenced == useNode ||
                (referenced.LocalName != "g" &&
                 referenced.LocalName != "symbol" &&
                 referenced.LocalName != "svg"))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(
                    GetAttr(
                        referenced,
                        "viewBox")) ||
                !string.IsNullOrEmpty(
                    GetAttr(
                        useNode,
                        "width")) ||
                !string.IsNullOrEmpty(
                    GetAttr(
                        useNode,
                        "height")))
            {
                return false;
            }

            string[] unsupportedRootEffects =
                new string[]
                {
                    "filter",
                    "clip-path",
                    "mask"
                };

            for (int i = 0;
                 i < unsupportedRootEffects.Length;
                 i++)
            {
                string key =
                    unsupportedRootEffects[i];

                if (!string.IsNullOrEmpty(
                        GetSvgStyle(
                            useNode,
                            key)) ||
                    !string.IsNullOrEmpty(
                        GetSvgStyle(
                            referenced,
                            key)))
                {
                    return false;
                }
            }

            XmlNode clone =
                referenced.CloneNode(
                    true);

            XmlElement cloneElement =
                clone as XmlElement;

            if (cloneElement == null)
            {
                return false;
            }

            cloneElement.SetAttribute(
                "data-ppl-use-expansion",
                "1");

            string[] inheritedKeys =
                new string[]
                {
                    "color",
                    "opacity",
                    "fill",
                    "fill-opacity",
                    "fill-rule",
                    "stroke",
                    "stroke-opacity",
                    "stroke-width",
                    "stroke-linecap",
                    "stroke-linejoin",
                    "stroke-miterlimit",
                    "stroke-dasharray",
                    "stroke-dashoffset"
                };

            for (int i = 0;
                 i < inheritedKeys.Length;
                 i++)
            {
                string key =
                    inheritedKeys[i];

                string existing =
                    GetSvgStyle(
                        clone,
                        key);

                if (!string.IsNullOrEmpty(
                        existing) &&
                    !string.Equals(
                        existing.Trim(),
                        "inherit",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string fallback =
                    GetSvgUseInvocationStyle(
                        useNode,
                        key);

                if (!string.IsNullOrEmpty(
                        fallback))
                {
                    cloneElement.SetAttribute(
                        key,
                        fallback);
                }
            }

            GraphicsState state =
                g.Save();

            try
            {
                ApplySimpleSvgTransform(
                    g,
                    useNode,
                    target,
                    minX,
                    minY,
                    sx,
                    sy);

                float useX =
                    ParseSvgFloat(
                        GetAttr(
                            useNode,
                            "x"),
                        0f) *
                    sx;
                float useY =
                    ParseSvgFloat(
                        GetAttr(
                            useNode,
                            "y"),
                        0f) *
                    sy;

                if (Math.Abs(useX) >
                        0.001f ||
                    Math.Abs(useY) >
                        0.001f)
                {
                    g.TranslateTransform(
                        useX,
                        useY);
                }

                ApplySimpleSvgTransform(
                    g,
                    clone,
                    target,
                    minX,
                    minY,
                    sx,
                    sy);

                return DrawEnhancedSvgChildren(
                    g,
                    clone,
                    target,
                    minX,
                    minY,
                    sx,
                    sy);
            }
            finally
            {
                g.Restore(
                    state);
            }
        }

        private static bool HasSvgUseExpansionMarker(
            XmlNode node)
        {
            XmlNode current =
                node;

            while (current != null)
            {
                if (string.Equals(
                        GetAttr(
                            current,
                            "data-ppl-use-expansion"),
                        "1",
                        StringComparison.Ordinal))
                {
                    return true;
                }

                current =
                    current.ParentNode;
            }

            return false;
        }

        private static string GetSvgUseInvocationStyle(
            XmlNode useNode,
            string key)
        {
            XmlNode current =
                useNode;

            while (current != null)
            {
                string value =
                    GetSvgStyle(
                        current,
                        key);

                if (!string.IsNullOrEmpty(
                        value))
                {
                    value =
                        value.Trim();

                    if (!string.Equals(
                            value,
                            "inherit",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return value;
                    }
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

            if (name == "use")
            {
                string referenceId =
                    ReadSvgUseReferenceId(
                        node);

                if (string.IsNullOrEmpty(
                        referenceId))
                {
                    return null;
                }

                XmlNode referenced =
                    FindSvgNodeById(
                        node.OwnerDocument,
                        referenceId);

                if (referenced == null ||
                    referenced == node)
                {
                    return null;
                }

                HashSet<string> activeReferences =
                    new HashSet<string>(
                        StringComparer.Ordinal);

                activeReferences.Add(
                    referenceId);

                GraphicsPath referencedPath =
                    BuildEnhancedSvgReferencePath(
                        referenced,
                        target,
                        minX,
                        minY,
                        sx,
                        sy,
                        activeReferences,
                        0);

                if (referencedPath == null ||
                    referencedPath.PointCount == 0)
                {
                    if (referencedPath != null)
                        referencedPath.Dispose();

                    return null;
                }

                ApplySvgUseViewportTransform(
                    referencedPath,
                    node,
                    referenced,
                    target,
                    minX,
                    minY,
                    sx,
                    sy);

                float useX =
                    ParseSvgFloat(
                        GetAttr(
                            node,
                            "x"),
                        0f) *
                    sx;
                float useY =
                    ParseSvgFloat(
                        GetAttr(
                            node,
                            "y"),
                        0f) *
                    sy;

                if (Math.Abs(useX) >
                        0.001f ||
                    Math.Abs(useY) >
                        0.001f)
                {
                    using (Matrix translation =
                        new Matrix())
                    {
                        translation.Translate(
                            useX,
                            useY);

                        referencedPath.Transform(
                            translation);
                    }
                }

                return referencedPath;
            }

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

        private static void ApplySvgUseViewportTransform(
            GraphicsPath path,
            XmlNode useNode,
            XmlNode referenced,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy)
        {
            if (path == null ||
                useNode == null ||
                referenced == null ||
                (referenced.LocalName != "symbol" &&
                 referenced.LocalName != "svg"))
            {
                return;
            }

            List<float> viewBox =
                ParseSvgNumberList(
                    GetAttr(
                        referenced,
                        "viewBox"));

            if (viewBox.Count < 4 ||
                Math.Abs(viewBox[2]) <
                    0.0001f ||
                Math.Abs(viewBox[3]) <
                    0.0001f)
            {
                return;
            }

            string widthText =
                GetAttr(
                    useNode,
                    "width");
            string heightText =
                GetAttr(
                    useNode,
                    "height");

            if (string.IsNullOrEmpty(
                    widthText) &&
                string.IsNullOrEmpty(
                    heightText))
            {
                return;
            }

            float viewportWidth =
                ParseSvgFloat(
                    widthText,
                    viewBox[2]);
            float viewportHeight =
                ParseSvgFloat(
                    heightText,
                    viewBox[3]);

            if (viewportWidth <= 0f ||
                viewportHeight <= 0f)
            {
                return;
            }

            float scaleX =
                viewportWidth /
                viewBox[2];
            float scaleY =
                viewportHeight /
                viewBox[3];

            float alignX = 0f;
            float alignY = 0f;

            string preserve =
                GetAttr(
                    useNode,
                    "preserveAspectRatio");

            if (string.IsNullOrEmpty(
                    preserve))
            {
                preserve =
                    GetAttr(
                        referenced,
                        "preserveAspectRatio");
            }

            if (string.IsNullOrEmpty(
                    preserve))
            {
                preserve =
                    "xMidYMid meet";
            }

            bool stretch =
                preserve.IndexOf(
                    "none",
                    StringComparison.OrdinalIgnoreCase) >=
                0;

            if (!stretch)
            {
                bool slice =
                    preserve.IndexOf(
                        "slice",
                        StringComparison.OrdinalIgnoreCase) >=
                    0;

                float uniform =
                    slice
                        ? Math.Max(
                            scaleX,
                            scaleY)
                        : Math.Min(
                            scaleX,
                            scaleY);

                float remainingX =
                    viewportWidth -
                    viewBox[2] *
                    uniform;
                float remainingY =
                    viewportHeight -
                    viewBox[3] *
                    uniform;

                if (preserve.IndexOf(
                        "xMax",
                        StringComparison.OrdinalIgnoreCase) >=
                    0)
                {
                    alignX =
                        remainingX *
                        sx;
                }
                else if (preserve.IndexOf(
                             "xMid",
                             StringComparison.OrdinalIgnoreCase) >=
                         0)
                {
                    alignX =
                        remainingX *
                        sx *
                        0.5f;
                }

                if (preserve.IndexOf(
                        "YMax",
                        StringComparison.OrdinalIgnoreCase) >=
                    0)
                {
                    alignY =
                        remainingY *
                        sy;
                }
                else if (preserve.IndexOf(
                             "YMid",
                             StringComparison.OrdinalIgnoreCase) >=
                         0)
                {
                    alignY =
                        remainingY *
                        sy *
                        0.5f;
                }

                scaleX =
                    uniform;
                scaleY =
                    uniform;
            }

            float originX =
                SvgX(
                    target,
                    minX,
                    sx,
                    viewBox[0]);
            float originY =
                SvgY(
                    target,
                    minY,
                    sy,
                    viewBox[1]);

            float offsetX =
                originX -
                originX *
                scaleX +
                alignX;
            float offsetY =
                originY -
                originY *
                scaleY +
                alignY;

            using (Matrix viewportMatrix =
                new Matrix(
                    scaleX,
                    0f,
                    0f,
                    scaleY,
                    offsetX,
                    offsetY))
            {
                path.Transform(
                    viewportMatrix);
            }
        }

        private static GraphicsPath BuildEnhancedSvgReferencePath(
            XmlNode node,
            RectangleF target,
            float minX,
            float minY,
            float sx,
            float sy,
            HashSet<string> activeReferences,
            int depth)
        {
            if (node == null ||
                depth > 12)
            {
                return null;
            }

            string name =
                node.LocalName;

            GraphicsPath result = null;

            if (name == "use")
            {
                string referenceId =
                    ReadSvgUseReferenceId(
                        node);

                if (string.IsNullOrEmpty(
                        referenceId) ||
                    activeReferences == null ||
                    activeReferences.Contains(
                        referenceId))
                {
                    return null;
                }

                XmlNode referenced =
                    FindSvgNodeById(
                        node.OwnerDocument,
                        referenceId);

                if (referenced == null ||
                    referenced == node)
                {
                    return null;
                }

                activeReferences.Add(
                    referenceId);

                try
                {
                    result =
                        BuildEnhancedSvgReferencePath(
                            referenced,
                            target,
                            minX,
                            minY,
                            sx,
                            sy,
                            activeReferences,
                            depth + 1);
                }
                finally
                {
                    activeReferences.Remove(
                        referenceId);
                }

                if (result == null ||
                    result.PointCount == 0)
                {
                    if (result != null)
                        result.Dispose();

                    return null;
                }

                ApplySvgUseViewportTransform(
                    result,
                    node,
                    referenced,
                    target,
                    minX,
                    minY,
                    sx,
                    sy);

                float useX =
                    ParseSvgFloat(
                        GetAttr(
                            node,
                            "x"),
                        0f) *
                    sx;
                float useY =
                    ParseSvgFloat(
                        GetAttr(
                            node,
                            "y"),
                        0f) *
                    sy;

                if (Math.Abs(useX) >
                        0.001f ||
                    Math.Abs(useY) >
                        0.001f)
                {
                    using (Matrix translation =
                        new Matrix())
                    {
                        translation.Translate(
                            useX,
                            useY);
                        result.Transform(
                            translation);
                    }
                }
            }
            else if (name == "g" ||
                     name == "symbol" ||
                     name == "svg")
            {
                result =
                    new GraphicsPath();

                for (int i = 0;
                     i < node.ChildNodes.Count;
                     i++)
                {
                    XmlNode child =
                        node.ChildNodes[i];

                    string childName =
                        child.LocalName;

                    if (childName == "defs" ||
                        childName == "linearGradient" ||
                        childName == "radialGradient" ||
                        childName == "clipPath" ||
                        childName == "pattern" ||
                        childName == "mask" ||
                        childName == "filter" ||
                        childName == "text")
                    {
                        continue;
                    }

                    GraphicsPath childPath = null;

                    try
                    {
                        if (childName == "g" ||
                            childName == "symbol" ||
                            childName == "svg" ||
                            childName == "use")
                        {
                            childPath =
                                BuildEnhancedSvgReferencePath(
                                    child,
                                    target,
                                    minX,
                                    minY,
                                    sx,
                                    sy,
                                    activeReferences,
                                    depth + 1);
                        }
                        else
                        {
                            childPath =
                                BuildEnhancedSvgElementPath(
                                    child,
                                    target,
                                    minX,
                                    minY,
                                    sx,
                                    sy);

                            if (childPath != null &&
                                childPath.PointCount > 0)
                            {
                                string childTransform =
                                    GetAttr(
                                        child,
                                        "transform");

                                if (!string.IsNullOrEmpty(
                                        childTransform))
                                {
                                    using (Matrix childMatrix =
                                        BuildSvgGeometryTransformMatrix(
                                            childTransform,
                                            target,
                                            minX,
                                            minY,
                                            sx,
                                            sy))
                                    {
                                        if (childMatrix != null &&
                                            !childMatrix.IsIdentity)
                                        {
                                            childPath.Transform(
                                                childMatrix);
                                        }
                                    }
                                }
                            }
                        }

                        if (childPath != null &&
                            childPath.PointCount > 0)
                        {
                            result.AddPath(
                                childPath,
                                false);
                        }
                    }
                    finally
                    {
                        if (childPath != null)
                            childPath.Dispose();
                    }
                }

                if (result.PointCount == 0)
                {
                    result.Dispose();
                    return null;
                }
            }
            else
            {
                result =
                    BuildEnhancedSvgElementPath(
                        node,
                        target,
                        minX,
                        minY,
                        sx,
                        sy);

                if (result == null ||
                    result.PointCount == 0)
                {
                    if (result != null)
                        result.Dispose();

                    return null;
                }
            }

            string transform =
                GetAttr(
                    node,
                    "transform");

            if (!string.IsNullOrEmpty(
                    transform))
            {
                using (Matrix matrix =
                    BuildSvgGeometryTransformMatrix(
                        transform,
                        target,
                        minX,
                        minY,
                        sx,
                        sy))
                {
                    if (matrix != null &&
                        !matrix.IsIdentity)
                    {
                        result.Transform(
                            matrix);
                    }
                }
            }

            return result;
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

        private static bool TryReadSvgColorMatrix(
            XmlNode node,
            XmlDocument document,
            out float[] matrix)
        {
            matrix = null;

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

            XmlNode colorMatrix =
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

                if (child.LocalName ==
                        "animate" ||
                    child.LocalName ==
                        "set")
                {
                    continue;
                }

                primitiveCount++;

                if (child.LocalName ==
                    "feColorMatrix")
                {
                    colorMatrix =
                        child;
                }
            }

            if (primitiveCount != 1 ||
                colorMatrix == null)
            {
                return false;
            }

            string input =
                GetAttr(
                    colorMatrix,
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

            string type =
                GetAttr(
                    colorMatrix,
                    "type");

            if (string.IsNullOrEmpty(
                    type))
            {
                type =
                    "matrix";
            }

            if (string.Equals(
                    type,
                    "matrix",
                    StringComparison.OrdinalIgnoreCase))
            {
                List<float> values =
                    ParseSvgNumberList(
                        GetAttr(
                            colorMatrix,
                            "values"));

                if (values.Count != 20)
                {
                    return false;
                }

                matrix =
                    values.ToArray();
                return true;
            }

            if (string.Equals(
                    type,
                    "saturate",
                    StringComparison.OrdinalIgnoreCase))
            {
                float saturation =
                    Math.Max(
                        0f,
                        ParseSvgFloat(
                            GetAttr(
                                colorMatrix,
                                "values"),
                            1f));

                matrix =
                    new float[]
                    {
                        0.213f + 0.787f * saturation,
                        0.715f - 0.715f * saturation,
                        0.072f - 0.072f * saturation,
                        0f,
                        0f,

                        0.213f - 0.213f * saturation,
                        0.715f + 0.285f * saturation,
                        0.072f - 0.072f * saturation,
                        0f,
                        0f,

                        0.213f - 0.213f * saturation,
                        0.715f - 0.715f * saturation,
                        0.072f + 0.928f * saturation,
                        0f,
                        0f,

                        0f,
                        0f,
                        0f,
                        1f,
                        0f
                    };

                return true;
            }

            if (string.Equals(
                    type,
                    "hueRotate",
                    StringComparison.OrdinalIgnoreCase))
            {
                float degrees =
                    ParseSvgFloat(
                        GetAttr(
                            colorMatrix,
                            "values"),
                        0f);

                double radians =
                    degrees *
                    Math.PI /
                    180.0;
                float cosine =
                    (float)Math.Cos(
                        radians);
                float sine =
                    (float)Math.Sin(
                        radians);

                matrix =
                    new float[]
                    {
                        0.213f + cosine * 0.787f - sine * 0.213f,
                        0.715f - cosine * 0.715f - sine * 0.715f,
                        0.072f - cosine * 0.072f + sine * 0.928f,
                        0f,
                        0f,

                        0.213f - cosine * 0.213f + sine * 0.143f,
                        0.715f + cosine * 0.285f + sine * 0.140f,
                        0.072f - cosine * 0.072f - sine * 0.283f,
                        0f,
                        0f,

                        0.213f - cosine * 0.213f - sine * 0.787f,
                        0.715f - cosine * 0.715f + sine * 0.715f,
                        0.072f + cosine * 0.928f + sine * 0.072f,
                        0f,
                        0f,

                        0f,
                        0f,
                        0f,
                        1f,
                        0f
                    };

                return true;
            }

            if (string.Equals(
                    type,
                    "luminanceToAlpha",
                    StringComparison.OrdinalIgnoreCase))
            {
                matrix =
                    new float[]
                    {
                        0f, 0f, 0f, 0f, 0f,
                        0f, 0f, 0f, 0f, 0f,
                        0f, 0f, 0f, 0f, 0f,
                        0.2125f, 0.7154f, 0.0721f, 0f, 0f
                    };

                return true;
            }

            return false;
        }

        private static Color ApplySvgColorMatrix(
            Color source,
            float[] matrix)
        {
            if (matrix == null ||
                matrix.Length != 20)
            {
                return source;
            }

            float r =
                source.R /
                255f;
            float g =
                source.G /
                255f;
            float b =
                source.B /
                255f;
            float a =
                source.A /
                255f;

            float outR =
                matrix[0] * r +
                matrix[1] * g +
                matrix[2] * b +
                matrix[3] * a +
                matrix[4];
            float outG =
                matrix[5] * r +
                matrix[6] * g +
                matrix[7] * b +
                matrix[8] * a +
                matrix[9];
            float outB =
                matrix[10] * r +
                matrix[11] * g +
                matrix[12] * b +
                matrix[13] * a +
                matrix[14];
            float outA =
                matrix[15] * r +
                matrix[16] * g +
                matrix[17] * b +
                matrix[18] * a +
                matrix[19];

            return Color.FromArgb(
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            Math.Max(
                                0f,
                                Math.Min(
                                    1f,
                                    outA)) *
                            255f))),
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            Math.Max(
                                0f,
                                Math.Min(
                                    1f,
                                    outR)) *
                            255f))),
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            Math.Max(
                                0f,
                                Math.Min(
                                    1f,
                                    outG)) *
                            255f))),
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            Math.Max(
                                0f,
                                Math.Min(
                                    1f,
                                    outB)) *
                            255f))));
        }

        private static bool DrawSvgColorMatrixApproximation(
            Graphics g,
            XmlNode node,
            XmlDocument document,
            GraphicsPath path,
            RectangleF target,
            float sx,
            float sy)
        {
            if (g == null ||
                node == null ||
                path == null ||
                path.PointCount == 0)
            {
                return false;
            }

            float[] matrix;

            if (!TryReadSvgColorMatrix(
                    node,
                    document,
                    out matrix))
            {
                return false;
            }

            string fillValue =
                GetSvgStyleInherited(
                    node,
                    "fill");

            bool hasFill =
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

            Color filteredFill =
                ApplySvgColorMatrix(
                    Color.FromArgb(
                        Math.Max(
                            0,
                            Math.Min(
                                255,
                                (int)Math.Round(
                                    fill.A *
                                    overallOpacity *
                                    fillOpacity))),
                        fill),
                    matrix);

            string strokeValue =
                GetSvgStyleInherited(
                    node,
                    "stroke");
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

            bool hasStroke =
                !string.IsNullOrEmpty(
                    strokeValue) &&
                !string.Equals(
                    strokeValue,
                    "none",
                    StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrEmpty(
                    ExtractSvgUrlId(
                        strokeValue));

            Color filteredStroke =
                ApplySvgColorMatrix(
                    Color.FromArgb(
                        Math.Max(
                            0,
                            Math.Min(
                                255,
                                (int)Math.Round(
                                    stroke.A *
                                    overallOpacity *
                                    strokeOpacity))),
                        stroke),
                    matrix);

            if ((!hasFill ||
                 filteredFill.A <= 0) &&
                (!hasStroke ||
                 filteredStroke.A <= 0))
            {
                return false;
            }

            if (hasFill &&
                filteredFill.A > 0)
            {
                using (Brush brush =
                    new SolidBrush(
                        filteredFill))
                {
                    g.FillPath(
                        brush,
                        path);
                }
            }

            if (hasStroke &&
                filteredStroke.A > 0)
            {
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

                using (Pen pen =
                    new Pen(
                        filteredStroke,
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
                        path);
                }
            }

            return true;
        }

        private static bool TryReadSvgOffsetGaussianChain(
            XmlNode node,
            XmlDocument document,
            float sx,
            float sy,
            out float offsetX,
            out float offsetY,
            out float blurX,
            out float blurY,
            out bool blendSourceGraphic,
            out string blendMode)
        {
            offsetX = 0f;
            offsetY = 0f;
            blurX = 0f;
            blurY = 0f;
            blendSourceGraphic = false;
            blendMode = string.Empty;

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

            List<XmlNode> primitives =
                new List<XmlNode>();

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

                if (child.LocalName ==
                        "animate" ||
                    child.LocalName ==
                        "set")
                {
                    continue;
                }

                primitives.Add(
                    child);
            }

            if (primitives.Count != 2 &&
                primitives.Count != 3)
            {
                return false;
            }

            XmlNode first =
                primitives[0];
            XmlNode second =
                primitives[1];

            XmlNode blend =
                primitives.Count == 3
                    ? primitives[2]
                    : null;

            bool offsetFirst =
                first.LocalName ==
                    "feOffset" &&
                second.LocalName ==
                    "feGaussianBlur";
            bool blurFirst =
                first.LocalName ==
                    "feGaussianBlur" &&
                second.LocalName ==
                    "feOffset";

            if (!offsetFirst &&
                !blurFirst)
            {
                return false;
            }

            XmlNode offset =
                offsetFirst
                    ? first
                    : second;
            XmlNode gaussian =
                offsetFirst
                    ? second
                    : first;

            string firstInput =
                GetAttr(
                    first,
                    "in");

            if (!string.IsNullOrEmpty(
                    firstInput) &&
                !string.Equals(
                    firstInput,
                    "SourceGraphic",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string firstResult =
                GetAttr(
                    first,
                    "result");
            string secondInput =
                GetAttr(
                    second,
                    "in");

            if (!string.IsNullOrEmpty(
                    secondInput))
            {
                if (string.IsNullOrEmpty(
                        firstResult) ||
                    !string.Equals(
                        secondInput,
                        firstResult,
                        StringComparison.Ordinal))
                {
                    return false;
                }
            }

            if (blend != null)
            {
                bool isBlend =
                    blend.LocalName ==
                    "feBlend";
                bool isComposite =
                    blend.LocalName ==
                    "feComposite";

                if (!isBlend &&
                    !isComposite)
                {
                    return false;
                }

                if (isBlend)
                {
                    string mode =
                        GetAttr(
                            blend,
                            "mode");

                    if (string.IsNullOrEmpty(
                            mode))
                    {
                        mode =
                            "normal";
                    }

                    if (!string.Equals(
                            mode,
                            "normal",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            mode,
                            "multiply",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            mode,
                            "screen",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            mode,
                            "darken",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            mode,
                            "lighten",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    blendMode =
                        mode.ToLowerInvariant();
                }
                else
                {
                    string op =
                        GetAttr(
                            blend,
                            "operator");

                    if (string.IsNullOrEmpty(
                            op))
                    {
                        op =
                            "over";
                    }

                    if (!string.Equals(
                            op,
                            "over",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            op,
                            "in",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            op,
                            "out",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            op,
                            "xor",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            op,
                            "atop",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(
                            op,
                            "arithmetic",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    blendMode =
                        op.ToLowerInvariant();
                }

                string secondResult =
                    GetAttr(
                        second,
                        "result");
                string blendIn =
                    GetAttr(
                        blend,
                        "in");
                string blendIn2 =
                    GetAttr(
                        blend,
                        "in2");

                bool firstIsChain =
                    !string.IsNullOrEmpty(
                        secondResult) &&
                    string.Equals(
                        blendIn,
                        secondResult,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        blendIn2,
                        "SourceGraphic",
                        StringComparison.OrdinalIgnoreCase);

                bool secondIsChain =
                    !string.IsNullOrEmpty(
                        secondResult) &&
                    string.Equals(
                        blendIn2,
                        secondResult,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        blendIn,
                        "SourceGraphic",
                        StringComparison.OrdinalIgnoreCase);

                if (!firstIsChain &&
                    !secondIsChain)
                {
                    return false;
                }

                if ((string.Equals(
                         blendMode,
                         "in",
                         StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(
                         blendMode,
                         "out",
                         StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(
                         blendMode,
                         "xor",
                         StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(
                         blendMode,
                         "atop",
                         StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(
                         blendMode,
                         "arithmetic",
                         StringComparison.OrdinalIgnoreCase)) &&
                    !secondIsChain)
                {
                    return false;
                }

                blendSourceGraphic = true;
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

            string rawDeviation =
                GetAttr(
                    gaussian,
                    "stdDeviation");

            if (string.IsNullOrEmpty(
                    rawDeviation))
            {
                return false;
            }

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

            return Math.Abs(offsetX) >
                    0.001f ||
                Math.Abs(offsetY) >
                    0.001f ||
                blurX > 0.05f ||
                blurY > 0.05f;
        }

        private static bool TryReadSvgCompositeArithmeticCoefficients(
            XmlNode node,
            XmlDocument document,
            out float k1,
            out float k2,
            out float k3,
            out float k4)
        {
            k1 = 0f;
            k2 = 0f;
            k3 = 0f;
            k4 = 0f;

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

            XmlNode arithmetic =
                null;

            for (int i = 0;
                 i < filter.ChildNodes.Count;
                 i++)
            {
                XmlNode child =
                    filter.ChildNodes[i];

                if (child == null ||
                    child.NodeType !=
                        XmlNodeType.Element ||
                    child.LocalName !=
                        "feComposite")
                {
                    continue;
                }

                if (string.Equals(
                        GetAttr(
                            child,
                            "operator"),
                        "arithmetic",
                        StringComparison.OrdinalIgnoreCase))
                {
                    arithmetic =
                        child;
                    break;
                }
            }

            if (arithmetic == null)
                return false;

            k1 =
                ClampSvgArithmeticCoefficient(
                    ParseSvgFloat(
                        GetAttr(
                            arithmetic,
                            "k1"),
                        0f));
            k2 =
                ClampSvgArithmeticCoefficient(
                    ParseSvgFloat(
                        GetAttr(
                            arithmetic,
                            "k2"),
                        0f));
            k3 =
                ClampSvgArithmeticCoefficient(
                    ParseSvgFloat(
                        GetAttr(
                            arithmetic,
                            "k3"),
                        0f));
            k4 =
                ClampSvgArithmeticCoefficient(
                    ParseSvgFloat(
                        GetAttr(
                            arithmetic,
                            "k4"),
                        0f));

            return true;
        }

        private static float ClampSvgArithmeticCoefficient(
            float value)
        {
            if (float.IsNaN(
                    value) ||
                float.IsInfinity(
                    value))
            {
                return 0f;
            }

            return Math.Max(
                -8f,
                Math.Min(
                    8f,
                    value));
        }

        private static bool DrawSvgOffsetGaussianChainApproximation(
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
                return false;
            }

            float offsetX;
            float offsetY;
            float blurX;
            float blurY;
            bool blendSourceGraphic;
            string blendMode;

            if (!TryReadSvgOffsetGaussianChain(
                    node,
                    document,
                    sx,
                    sy,
                    out offsetX,
                    out offsetY,
                    out blurX,
                    out blurY,
                    out blendSourceGraphic,
                    out blendMode))
            {
                return false;
            }

            string fillValue =
                GetSvgStyleInherited(
                    node,
                    "fill");

            bool hasSolidFill =
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

            if ((!hasSolidFill ||
                 fillAlpha <= 0) &&
                strokeAlpha <= 0)
            {
                return false;
            }

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

            if (blendSourceGraphic &&
                string.Equals(
                    blendMode,
                    "arithmetic",
                    StringComparison.OrdinalIgnoreCase))
            {
                float k1;
                float k2;
                float k3;
                float k4;

                if (!TryReadSvgCompositeArithmeticCoefficients(
                        node,
                        document,
                        out k1,
                        out k2,
                        out k3,
                        out k4))
                {
                    return false;
                }

                DrawSvgCompositeArithmeticApproximation(
                    g,
                    path,
                    offsetX,
                    offsetY,
                    hasSolidFill,
                    fillAlpha,
                    fill,
                    strokeAlpha,
                    stroke,
                    strokeWidth,
                    k1,
                    k2,
                    k3,
                    k4);
                return true;
            }

            if (blendSourceGraphic &&
                string.Equals(
                    blendMode,
                    "in",
                    StringComparison.OrdinalIgnoreCase))
            {
                DrawSvgCompositeInApproximation(
                    g,
                    path,
                    offsetX,
                    offsetY,
                    hasSolidFill,
                    fillAlpha,
                    fill,
                    strokeAlpha,
                    stroke,
                    strokeWidth);
                return true;
            }

            if (blendSourceGraphic &&
                string.Equals(
                    blendMode,
                    "out",
                    StringComparison.OrdinalIgnoreCase))
            {
                DrawSvgCompositeOutApproximation(
                    g,
                    path,
                    offsetX,
                    offsetY,
                    hasSolidFill,
                    fillAlpha,
                    fill,
                    strokeAlpha,
                    stroke,
                    strokeWidth);
                return true;
            }

            if (blendSourceGraphic &&
                string.Equals(
                    blendMode,
                    "xor",
                    StringComparison.OrdinalIgnoreCase))
            {
                DrawSvgCompositeXorApproximation(
                    g,
                    path,
                    offsetX,
                    offsetY,
                    hasSolidFill,
                    fillAlpha,
                    fill,
                    strokeAlpha,
                    stroke,
                    strokeWidth);
                return true;
            }

            if (blendSourceGraphic &&
                string.Equals(
                    blendMode,
                    "atop",
                    StringComparison.OrdinalIgnoreCase))
            {
                DrawSvgCompositeAtopApproximation(
                    g,
                    path,
                    offsetX,
                    offsetY,
                    hasSolidFill,
                    fillAlpha,
                    fill,
                    strokeAlpha,
                    stroke,
                    strokeWidth);
                return true;
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

                    if (hasSolidFill &&
                        fillAlpha > 0)
                    {
                        using (Brush brush =
                            new SolidBrush(
                                Color.FromArgb(
                                    fillAlpha,
                                    fill)))
                        {
                            g.FillPath(
                                brush,
                                shifted);
                        }
                    }

                    if (strokeAlpha > 0)
                    {
                        using (Pen pen =
                            new Pen(
                                Color.FromArgb(
                                    strokeAlpha,
                                    stroke),
                                strokeWidth))
                        {
                            g.DrawPath(
                                pen,
                                shifted);
                        }
                    }
                }

                if (blendSourceGraphic)
                {
                    if (hasSolidFill &&
                        fillAlpha > 0)
                    {
                        using (Brush sourceBrush =
                            new SolidBrush(
                                Color.FromArgb(
                                    fillAlpha,
                                    fill)))
                        {
                            g.FillPath(
                                sourceBrush,
                                path);
                        }
                    }

                    if (strokeAlpha > 0)
                    {
                        using (Pen sourcePen =
                            new Pen(
                                Color.FromArgb(
                                    strokeAlpha,
                                    stroke),
                                strokeWidth))
                        {
                            g.DrawPath(
                                sourcePen,
                                path);
                        }
                    }

                    if (string.Equals(
                            blendMode,
                            "multiply",
                            StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(
                            blendMode,
                            "screen",
                            StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(
                            blendMode,
                            "darken",
                            StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(
                            blendMode,
                            "lighten",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        DrawSvgBlendOverlapApproximation(
                            g,
                            path,
                            offsetX,
                            offsetY,
                            hasSolidFill,
                            fillAlpha,
                            fill,
                            strokeAlpha,
                            stroke,
                            strokeWidth,
                            blendMode);
                    }
                }

                return true;
            }

            const int rings = 6;
            const int directions = 16;

            for (int ring = 0;
                 ring <= rings;
                 ring++)
            {
                float ringRatio =
                    ring /
                    (float)rings;

                int directionCount =
                    ring == 0
                        ? 1
                        : directions;

                double sigmaDistance =
                    2.6 *
                    ringRatio;

                float gaussianWeight =
                    (float)Math.Exp(
                        -0.5 *
                        sigmaDistance *
                        sigmaDistance);

                float alphaRatio =
                    ring == 0
                        ? 0.30f
                        : Math.Max(
                            0.004f,
                            gaussianWeight /
                            (directions *
                             0.72f));

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

                    float blurOffsetX =
                        ring == 0
                            ? 0f
                            : (float)Math.Cos(
                                angle) *
                              blurX *
                              3f *
                              ringRatio;
                    float blurOffsetY =
                        ring == 0
                            ? 0f
                            : (float)Math.Sin(
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
                            offsetX +
                            blurOffsetX,
                            offsetY +
                            blurOffsetY);
                        shifted.Transform(
                            translation);

                        if (hasSolidFill &&
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

            if (blendSourceGraphic)
            {
                if (hasSolidFill &&
                    fillAlpha > 0)
                {
                    using (Brush sourceBrush =
                        new SolidBrush(
                            Color.FromArgb(
                                fillAlpha,
                                fill)))
                    {
                        g.FillPath(
                            sourceBrush,
                            path);
                    }
                }

                if (strokeAlpha > 0)
                {
                    using (Pen sourcePen =
                        new Pen(
                            Color.FromArgb(
                                strokeAlpha,
                                stroke),
                            strokeWidth))
                    {
                        g.DrawPath(
                            sourcePen,
                            path);
                    }
                }

                if (string.Equals(
                        blendMode,
                        "multiply",
                        StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(
                        blendMode,
                        "screen",
                        StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(
                        blendMode,
                        "darken",
                        StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(
                        blendMode,
                        "lighten",
                        StringComparison.OrdinalIgnoreCase))
                {
                    DrawSvgBlendOverlapApproximation(
                        g,
                        path,
                        offsetX,
                        offsetY,
                        hasSolidFill,
                        fillAlpha,
                        fill,
                        strokeAlpha,
                        stroke,
                        strokeWidth,
                        blendMode);
                }
            }

            return true;
        }

        private static void DrawSvgCompositeArithmeticApproximation(
            Graphics g,
            GraphicsPath sourcePath,
            float offsetX,
            float offsetY,
            bool hasSolidFill,
            int fillAlpha,
            Color fill,
            int strokeAlpha,
            Color stroke,
            float strokeWidth,
            float k1,
            float k2,
            float k3,
            float k4)
        {
            if (g == null ||
                sourcePath == null ||
                sourcePath.PointCount == 0)
            {
                return;
            }

            using (GraphicsPath shifted =
                (GraphicsPath)sourcePath.Clone())
            using (Matrix translation =
                new Matrix())
            {
                translation.Translate(
                    offsetX,
                    offsetY);
                shifted.Transform(
                    translation);

                using (Region sourceOnly =
                    new Region(
                        sourcePath))
                using (Region shiftedOnly =
                    new Region(
                        shifted))
                using (Region overlap =
                    new Region(
                        sourcePath))
                {
                    sourceOnly.Exclude(
                        shifted);
                    shiftedOnly.Exclude(
                        sourcePath);
                    overlap.Intersect(
                        shifted);

                    DrawSvgArithmeticRegion(
                        g,
                        sourceOnly,
                        sourcePath,
                        true,
                        false,
                        hasSolidFill,
                        fillAlpha,
                        fill,
                        strokeAlpha,
                        stroke,
                        strokeWidth,
                        k1,
                        k2,
                        k3,
                        k4);

                    DrawSvgArithmeticRegion(
                        g,
                        overlap,
                        sourcePath,
                        true,
                        true,
                        hasSolidFill,
                        fillAlpha,
                        fill,
                        strokeAlpha,
                        stroke,
                        strokeWidth,
                        k1,
                        k2,
                        k3,
                        k4);

                    DrawSvgArithmeticRegion(
                        g,
                        shiftedOnly,
                        shifted,
                        false,
                        true,
                        hasSolidFill,
                        fillAlpha,
                        fill,
                        strokeAlpha,
                        stroke,
                        strokeWidth,
                        k1,
                        k2,
                        k3,
                        k4);
                }
            }
        }

        private static void DrawSvgArithmeticRegion(
            Graphics g,
            Region clip,
            GraphicsPath renderPath,
            bool sourcePresent,
            bool shiftedPresent,
            bool hasSolidFill,
            int fillAlpha,
            Color fill,
            int strokeAlpha,
            Color stroke,
            float strokeWidth,
            float k1,
            float k2,
            float k3,
            float k4)
        {
            if (g == null ||
                clip == null ||
                renderPath == null)
            {
                return;
            }

            GraphicsState state =
                g.Save();

            try
            {
                g.SetClip(
                    clip,
                    CombineMode.Intersect);

                if (hasSolidFill &&
                    fillAlpha > 0)
                {
                    Color arithmeticFill =
                        EvaluateSvgArithmeticColor(
                            fill,
                            fillAlpha,
                            sourcePresent,
                            shiftedPresent,
                            k1,
                            k2,
                            k3,
                            k4);

                    if (arithmeticFill.A > 0)
                    {
                        using (Brush brush =
                            new SolidBrush(
                                arithmeticFill))
                        {
                            g.FillPath(
                                brush,
                                renderPath);
                        }
                    }
                }

                if (strokeAlpha > 0)
                {
                    Color arithmeticStroke =
                        EvaluateSvgArithmeticColor(
                            stroke,
                            strokeAlpha,
                            sourcePresent,
                            shiftedPresent,
                            k1,
                            k2,
                            k3,
                            k4);

                    if (arithmeticStroke.A > 0)
                    {
                        using (Pen pen =
                            new Pen(
                                arithmeticStroke,
                                strokeWidth))
                        {
                            g.DrawPath(
                                pen,
                                renderPath);
                        }
                    }
                }
            }
            finally
            {
                g.Restore(
                    state);
            }
        }

        private static Color EvaluateSvgArithmeticColor(
            Color color,
            int alpha,
            bool sourcePresent,
            bool shiftedPresent,
            float k1,
            float k2,
            float k3,
            float k4)
        {
            float sourceAlpha =
                sourcePresent
                    ? Math.Max(
                        0f,
                        Math.Min(
                            1f,
                            alpha /
                            255f))
                    : 0f;
            float shiftedAlpha =
                shiftedPresent
                    ? Math.Max(
                        0f,
                        Math.Min(
                            1f,
                            alpha /
                            255f))
                    : 0f;

            float resultAlpha =
                EvaluateSvgArithmeticChannel(
                    sourceAlpha,
                    shiftedAlpha,
                    k1,
                    k2,
                    k3,
                    k4);

            float red =
                EvaluateSvgArithmeticChannel(
                    sourcePresent
                        ? color.R /
                          255f
                        : 0f,
                    shiftedPresent
                        ? color.R /
                          255f
                        : 0f,
                    k1,
                    k2,
                    k3,
                    k4);
            float green =
                EvaluateSvgArithmeticChannel(
                    sourcePresent
                        ? color.G /
                          255f
                        : 0f,
                    shiftedPresent
                        ? color.G /
                          255f
                        : 0f,
                    k1,
                    k2,
                    k3,
                    k4);
            float blue =
                EvaluateSvgArithmeticChannel(
                    sourcePresent
                        ? color.B /
                          255f
                        : 0f,
                    shiftedPresent
                        ? color.B /
                          255f
                        : 0f,
                    k1,
                    k2,
                    k3,
                    k4);

            return Color.FromArgb(
                (int)Math.Round(
                    resultAlpha *
                    255f),
                (int)Math.Round(
                    red *
                    255f),
                (int)Math.Round(
                    green *
                    255f),
                (int)Math.Round(
                    blue *
                    255f));
        }

        private static float EvaluateSvgArithmeticChannel(
            float first,
            float second,
            float k1,
            float k2,
            float k3,
            float k4)
        {
            float value =
                k1 *
                first *
                second +
                k2 *
                first +
                k3 *
                second +
                k4;

            return Math.Max(
                0f,
                Math.Min(
                    1f,
                    value));
        }

        private static void DrawSvgCompositeInApproximation(
            Graphics g,
            GraphicsPath sourcePath,
            float offsetX,
            float offsetY,
            bool hasSolidFill,
            int fillAlpha,
            Color fill,
            int strokeAlpha,
            Color stroke,
            float strokeWidth)
        {
            if (g == null ||
                sourcePath == null ||
                sourcePath.PointCount == 0)
            {
                return;
            }

            using (GraphicsPath shifted =
                (GraphicsPath)sourcePath.Clone())
            using (Matrix translation =
                new Matrix())
            {
                translation.Translate(
                    offsetX,
                    offsetY);
                shifted.Transform(
                    translation);

                using (Region overlap =
                    new Region(
                        sourcePath))
                {
                    overlap.Intersect(
                        shifted);

                    GraphicsState state =
                        g.Save();

                    try
                    {
                        g.SetClip(
                            overlap,
                            CombineMode.Intersect);

                        if (hasSolidFill &&
                            fillAlpha > 0)
                        {
                            using (Brush brush =
                                new SolidBrush(
                                    Color.FromArgb(
                                        fillAlpha,
                                        fill)))
                            {
                                g.FillPath(
                                    brush,
                                    sourcePath);
                            }
                        }

                        if (strokeAlpha > 0)
                        {
                            using (Pen pen =
                                new Pen(
                                    Color.FromArgb(
                                        strokeAlpha,
                                        stroke),
                                    strokeWidth))
                            {
                                g.DrawPath(
                                    pen,
                                    sourcePath);
                            }
                        }
                    }
                    finally
                    {
                        g.Restore(
                            state);
                    }
                }
            }
        }

        private static void DrawSvgCompositeOutApproximation(
            Graphics g,
            GraphicsPath sourcePath,
            float offsetX,
            float offsetY,
            bool hasSolidFill,
            int fillAlpha,
            Color fill,
            int strokeAlpha,
            Color stroke,
            float strokeWidth)
        {
            if (g == null ||
                sourcePath == null ||
                sourcePath.PointCount == 0)
            {
                return;
            }

            using (GraphicsPath shifted =
                (GraphicsPath)sourcePath.Clone())
            using (Matrix translation =
                new Matrix())
            {
                translation.Translate(
                    offsetX,
                    offsetY);
                shifted.Transform(
                    translation);

                using (Region outside =
                    new Region(
                        sourcePath))
                {
                    outside.Exclude(
                        shifted);

                    GraphicsState state =
                        g.Save();

                    try
                    {
                        g.SetClip(
                            outside,
                            CombineMode.Intersect);

                        if (hasSolidFill &&
                            fillAlpha > 0)
                        {
                            using (Brush brush =
                                new SolidBrush(
                                    Color.FromArgb(
                                        fillAlpha,
                                        fill)))
                            {
                                g.FillPath(
                                    brush,
                                    sourcePath);
                            }
                        }

                        if (strokeAlpha > 0)
                        {
                            using (Pen pen =
                                new Pen(
                                    Color.FromArgb(
                                        strokeAlpha,
                                        stroke),
                                    strokeWidth))
                            {
                                g.DrawPath(
                                    pen,
                                    sourcePath);
                            }
                        }
                    }
                    finally
                    {
                        g.Restore(
                            state);
                    }
                }
            }
        }

        private static void DrawSvgCompositeXorApproximation(
            Graphics g,
            GraphicsPath sourcePath,
            float offsetX,
            float offsetY,
            bool hasSolidFill,
            int fillAlpha,
            Color fill,
            int strokeAlpha,
            Color stroke,
            float strokeWidth)
        {
            if (g == null ||
                sourcePath == null ||
                sourcePath.PointCount == 0)
            {
                return;
            }

            using (GraphicsPath shifted =
                (GraphicsPath)sourcePath.Clone())
            using (Matrix translation =
                new Matrix())
            {
                translation.Translate(
                    offsetX,
                    offsetY);
                shifted.Transform(
                    translation);

                using (Region sourceOnly =
                    new Region(
                        sourcePath))
                using (Region shiftedOnly =
                    new Region(
                        shifted))
                {
                    sourceOnly.Exclude(
                        shifted);
                    shiftedOnly.Exclude(
                        sourcePath);

                    GraphicsState state =
                        g.Save();

                    try
                    {
                        g.SetClip(
                            sourceOnly,
                            CombineMode.Intersect);

                        if (hasSolidFill &&
                            fillAlpha > 0)
                        {
                            using (Brush brush =
                                new SolidBrush(
                                    Color.FromArgb(
                                        fillAlpha,
                                        fill)))
                            {
                                g.FillPath(
                                    brush,
                                    sourcePath);
                            }
                        }

                        if (strokeAlpha > 0)
                        {
                            using (Pen pen =
                                new Pen(
                                    Color.FromArgb(
                                        strokeAlpha,
                                        stroke),
                                    strokeWidth))
                            {
                                g.DrawPath(
                                    pen,
                                    sourcePath);
                            }
                        }
                    }
                    finally
                    {
                        g.Restore(
                            state);
                    }

                    state =
                        g.Save();

                    try
                    {
                        g.SetClip(
                            shiftedOnly,
                            CombineMode.Intersect);

                        if (hasSolidFill &&
                            fillAlpha > 0)
                        {
                            using (Brush brush =
                                new SolidBrush(
                                    Color.FromArgb(
                                        fillAlpha,
                                        fill)))
                            {
                                g.FillPath(
                                    brush,
                                    shifted);
                            }
                        }

                        if (strokeAlpha > 0)
                        {
                            using (Pen pen =
                                new Pen(
                                    Color.FromArgb(
                                        strokeAlpha,
                                        stroke),
                                    strokeWidth))
                            {
                                g.DrawPath(
                                    pen,
                                    shifted);
                            }
                        }
                    }
                    finally
                    {
                        g.Restore(
                            state);
                    }
                }
            }
        }

        private static void DrawSvgCompositeAtopApproximation(
            Graphics g,
            GraphicsPath sourcePath,
            float offsetX,
            float offsetY,
            bool hasSolidFill,
            int fillAlpha,
            Color fill,
            int strokeAlpha,
            Color stroke,
            float strokeWidth)
        {
            if (g == null ||
                sourcePath == null ||
                sourcePath.PointCount == 0)
            {
                return;
            }

            using (GraphicsPath shifted =
                (GraphicsPath)sourcePath.Clone())
            using (Matrix translation =
                new Matrix())
            {
                translation.Translate(
                    offsetX,
                    offsetY);
                shifted.Transform(
                    translation);

                if (hasSolidFill &&
                    fillAlpha > 0)
                {
                    using (Brush brush =
                        new SolidBrush(
                            Color.FromArgb(
                                fillAlpha,
                                fill)))
                    {
                        g.FillPath(
                            brush,
                            shifted);
                    }
                }

                if (strokeAlpha > 0)
                {
                    using (Pen pen =
                        new Pen(
                            Color.FromArgb(
                                strokeAlpha,
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

        private static void DrawSvgBlendOverlapApproximation(
            Graphics g,
            GraphicsPath sourcePath,
            float offsetX,
            float offsetY,
            bool hasSolidFill,
            int fillAlpha,
            Color fill,
            int strokeAlpha,
            Color stroke,
            float strokeWidth,
            string blendMode)
        {
            if (g == null ||
                sourcePath == null ||
                sourcePath.PointCount == 0)
            {
                return;
            }

            using (GraphicsPath shifted =
                (GraphicsPath)sourcePath.Clone())
            using (Matrix translation =
                new Matrix())
            {
                translation.Translate(
                    offsetX,
                    offsetY);
                shifted.Transform(
                    translation);

                using (Region overlap =
                    new Region(
                        sourcePath))
                {
                    overlap.Intersect(
                        shifted);

                    GraphicsState state =
                        g.Save();

                    try
                    {
                        g.SetClip(
                            overlap,
                            CombineMode.Intersect);

                        if (hasSolidFill &&
                            fillAlpha > 0)
                        {
                            Color blendedFill =
                                Color.FromArgb(
                                    fillAlpha,
                                    BlendSvgChannel(
                                        fill.R,
                                        blendMode),
                                    BlendSvgChannel(
                                        fill.G,
                                        blendMode),
                                    BlendSvgChannel(
                                        fill.B,
                                        blendMode));

                            using (Brush brush =
                                new SolidBrush(
                                    blendedFill))
                            {
                                g.FillPath(
                                    brush,
                                    sourcePath);
                            }
                        }

                        if (strokeAlpha > 0)
                        {
                            Color blendedStroke =
                                Color.FromArgb(
                                    strokeAlpha,
                                    BlendSvgChannel(
                                        stroke.R,
                                        blendMode),
                                    BlendSvgChannel(
                                        stroke.G,
                                        blendMode),
                                    BlendSvgChannel(
                                        stroke.B,
                                        blendMode));

                            using (Pen pen =
                                new Pen(
                                    blendedStroke,
                                    strokeWidth))
                            {
                                g.DrawPath(
                                    pen,
                                    sourcePath);
                            }
                        }
                    }
                    finally
                    {
                        g.Restore(
                            state);
                    }
                }
            }
        }

        private static int BlendSvgChannel(
            int channel,
            string blendMode)
        {
            channel =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        channel));

            if (string.Equals(
                    blendMode,
                    "screen",
                    StringComparison.OrdinalIgnoreCase))
            {
                int inverse =
                    255 -
                    channel;

                return 255 -
                    inverse *
                    inverse /
                    255;
            }

            if (string.Equals(
                    blendMode,
                    "darken",
                    StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    blendMode,
                    "lighten",
                    StringComparison.OrdinalIgnoreCase))
            {
                return channel;
            }

            return channel *
                channel /
                255;
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

        private static string ReadSvgUseReferenceId(
            XmlNode node)
        {
            if (node == null ||
                node.Attributes == null)
            {
                return null;
            }

            string raw =
                null;

            for (int i = 0;
                 i < node.Attributes.Count;
                 i++)
            {
                XmlAttribute attribute =
                    node.Attributes[i];

                if (attribute != null &&
                    string.Equals(
                        attribute.LocalName,
                        "href",
                        StringComparison.OrdinalIgnoreCase))
                {
                    raw =
                        attribute.Value;
                    break;
                }
            }

            if (string.IsNullOrEmpty(raw))
            {
                return null;
            }

            raw =
                raw.Trim();

            if (raw.StartsWith(
                    "#",
                    StringComparison.Ordinal))
            {
                raw =
                    raw.Substring(1);
            }

            return raw.Length > 0
                ? raw
                : null;
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
            if (node != null &&
                node.LocalName == "use")
            {
                string referencedValue =
                    GetSvgUseReferencedStyle(
                        node,
                        key,
                        new HashSet<string>(
                            StringComparer.Ordinal),
                        0);

                if (!string.IsNullOrEmpty(
                        referencedValue))
                {
                    return referencedValue;
                }
            }

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
                    value =
                        value.Trim();

                    if (!string.Equals(
                            value,
                            "inherit",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return value;
                    }
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

        private static string GetSvgUseReferencedStyle(
            XmlNode useNode,
            string key,
            HashSet<string> activeReferences,
            int depth)
        {
            if (useNode == null ||
                useNode.LocalName != "use" ||
                activeReferences == null ||
                depth > 12)
            {
                return null;
            }

            string referenceId =
                ReadSvgUseReferenceId(
                    useNode);

            if (string.IsNullOrEmpty(
                    referenceId) ||
                activeReferences.Contains(
                    referenceId))
            {
                return null;
            }

            XmlNode referenced =
                FindSvgNodeById(
                    useNode.OwnerDocument,
                    referenceId);

            if (referenced == null ||
                referenced == useNode)
            {
                return null;
            }

            activeReferences.Add(
                referenceId);

            try
            {
                string referencedValue =
                    GetSvgStyle(
                        referenced,
                        key);

                if (!string.IsNullOrEmpty(
                        referencedValue))
                {
                    referencedValue =
                        referencedValue.Trim();

                    if (!string.Equals(
                            referencedValue,
                            "inherit",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return referencedValue;
                    }
                }

                if (referenced.LocalName ==
                    "use")
                {
                    return GetSvgUseReferencedStyle(
                        referenced,
                        key,
                        activeReferences,
                        depth + 1);
                }

                return null;
            }
            finally
            {
                activeReferences.Remove(
                    referenceId);
            }
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

            if (string.Equals(
                    raw,
                    "currentColor",
                    StringComparison.OrdinalIgnoreCase))
            {
                string current =
                    GetSvgStyleInherited(
                        node,
                        "color");

                if (string.IsNullOrEmpty(current) ||
                    string.Equals(
                        current.Trim(),
                        "currentColor",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return Color.Black;
                }

                return ParseSvgColorValue(
                    current,
                    Color.Black);
            }

            return ParseSvgColorValue(
                raw,
                fallback);
        }

        private static Color ParseSvgColorValue(
            string raw,
            Color fallback)
        {
            if (string.IsNullOrEmpty(raw))
                return fallback;

            raw = raw.Trim();

            if (string.Equals(
                    raw,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                return Color.Transparent;
            }

            Color hexColor;

            if (TryParseSvgHexColor(
                    raw,
                    out hexColor))
            {
                return hexColor;
            }

            Color rgbColor;

            if (TryParseSvgRgbColor(
                    raw,
                    out rgbColor))
            {
                return rgbColor;
            }

            Color hslColor;

            if (TryParseSvgHslColor(
                    raw,
                    out hslColor))
            {
                return hslColor;
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

        private static bool TryParseSvgHexColor(
            string raw,
            out Color color)
        {
            color =
                Color.Empty;

            if (string.IsNullOrEmpty(raw))
                return false;

            raw =
                raw.Trim();

            if (!raw.StartsWith(
                    "#",
                    StringComparison.Ordinal))
            {
                return false;
            }

            string hex =
                raw.Substring(1);

            if (hex.Length == 3 ||
                hex.Length == 4)
            {
                StringBuilder expanded =
                    new StringBuilder();

                for (int i = 0;
                     i < hex.Length;
                     i++)
                {
                    expanded.Append(
                        hex[i]);
                    expanded.Append(
                        hex[i]);
                }

                hex =
                    expanded.ToString();
            }

            if (hex.Length != 6 &&
                hex.Length != 8)
            {
                return false;
            }

            int red;
            int green;
            int blue;
            int alpha =
                255;

            if (!int.TryParse(
                    hex.Substring(
                        0,
                        2),
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out red) ||
                !int.TryParse(
                    hex.Substring(
                        2,
                        2),
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out green) ||
                !int.TryParse(
                    hex.Substring(
                        4,
                        2),
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out blue))
            {
                return false;
            }

            if (hex.Length == 8 &&
                !int.TryParse(
                    hex.Substring(
                        6,
                        2),
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out alpha))
            {
                return false;
            }

            color =
                Color.FromArgb(
                    alpha,
                    red,
                    green,
                    blue);

            return true;
        }

        private static bool TryParseSvgRgbColor(
            string raw,
            out Color color)
        {
            color =
                Color.Empty;

            if (string.IsNullOrEmpty(raw))
            {
                return false;
            }

            raw =
                raw.Trim();

            bool rgb =
                raw.StartsWith(
                    "rgb(",
                    StringComparison.OrdinalIgnoreCase);
            bool rgba =
                raw.StartsWith(
                    "rgba(",
                    StringComparison.OrdinalIgnoreCase);

            if (!rgb &&
                !rgba)
            {
                return false;
            }

            int open =
                raw.IndexOf('(');
            int close =
                raw.LastIndexOf(')');

            if (open < 0 ||
                close <= open)
            {
                return false;
            }

            string body =
                raw.Substring(
                    open + 1,
                    close -
                    open -
                    1)
                .Trim();

            if (body.StartsWith(
                    "from ",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            body =
                body.Replace(
                    ",",
                    " ")
                .Replace(
                    "/",
                    " / ");

            string[] tokens =
                body.Split(
                    new char[]
                    {
                        ' ',
                        '\t',
                        '\r',
                        '\n'
                    },
                    StringSplitOptions.RemoveEmptyEntries);

            List<string> components =
                new List<string>();
            string alphaToken =
                null;
            bool afterSlash =
                false;

            for (int i = 0;
                 i < tokens.Length;
                 i++)
            {
                string token =
                    tokens[i];

                if (token == "/")
                {
                    afterSlash =
                        true;
                    continue;
                }

                if (afterSlash)
                {
                    if (alphaToken == null)
                    {
                        alphaToken =
                            token;
                    }

                    continue;
                }

                components.Add(
                    token);
            }

            if (components.Count == 4 &&
                alphaToken == null)
            {
                alphaToken =
                    components[3];
                components.RemoveAt(3);
            }

            if (components.Count != 3)
            {
                return false;
            }

            int red;
            int green;
            int blue;

            if (!TryParseSvgRgbChannel(
                    components[0],
                    out red) ||
                !TryParseSvgRgbChannel(
                    components[1],
                    out green) ||
                !TryParseSvgRgbChannel(
                    components[2],
                    out blue))
            {
                return false;
            }

            int alpha =
                255;

            if (!string.IsNullOrEmpty(
                    alphaToken) &&
                !TryParseSvgAlphaChannel(
                    alphaToken,
                    out alpha))
            {
                return false;
            }

            color =
                Color.FromArgb(
                    alpha,
                    red,
                    green,
                    blue);

            return true;
        }

        private static bool TryParseSvgRgbChannel(
            string raw,
            out int value)
        {
            value = 0;

            if (string.IsNullOrEmpty(raw))
            {
                return false;
            }

            raw =
                raw.Trim();

            if (string.Equals(
                    raw,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                value = 0;
                return true;
            }

            bool percent =
                raw.EndsWith(
                    "%",
                    StringComparison.Ordinal);

            string number =
                percent
                    ? raw.Substring(
                        0,
                        raw.Length - 1)
                    : raw;

            float parsed =
                ParseSvgFloat(
                    number,
                    float.NaN);

            if (float.IsNaN(parsed) ||
                float.IsInfinity(parsed))
            {
                return false;
            }

            float channel =
                percent
                    ? parsed *
                      255f /
                      100f
                    : parsed;

            value =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            channel)));

            return true;
        }

        private static bool TryParseSvgHslColor(
            string raw,
            out Color color)
        {
            color =
                Color.Empty;

            if (string.IsNullOrEmpty(raw))
                return false;

            raw =
                raw.Trim();

            bool hsl =
                raw.StartsWith(
                    "hsl(",
                    StringComparison.OrdinalIgnoreCase);
            bool hsla =
                raw.StartsWith(
                    "hsla(",
                    StringComparison.OrdinalIgnoreCase);

            if (!hsl &&
                !hsla)
            {
                return false;
            }

            int open =
                raw.IndexOf('(');
            int close =
                raw.LastIndexOf(')');

            if (open < 0 ||
                close <= open)
            {
                return false;
            }

            string body =
                raw.Substring(
                    open + 1,
                    close -
                    open -
                    1)
                .Trim();

            if (body.StartsWith(
                    "from ",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            body =
                body.Replace(
                    ",",
                    " ")
                .Replace(
                    "/",
                    " / ");

            string[] tokens =
                body.Split(
                    new char[]
                    {
                        ' ',
                        '\t',
                        '\r',
                        '\n'
                    },
                    StringSplitOptions.RemoveEmptyEntries);

            List<string> components =
                new List<string>();
            string alphaToken =
                null;
            bool afterSlash =
                false;

            for (int i = 0;
                 i < tokens.Length;
                 i++)
            {
                string token =
                    tokens[i];

                if (token == "/")
                {
                    afterSlash =
                        true;
                    continue;
                }

                if (afterSlash)
                {
                    if (alphaToken == null)
                    {
                        alphaToken =
                            token;
                    }

                    continue;
                }

                components.Add(
                    token);
            }

            if (components.Count == 4 &&
                alphaToken == null)
            {
                alphaToken =
                    components[3];
                components.RemoveAt(3);
            }

            if (components.Count != 3)
                return false;

            float hue;
            float saturation;
            float lightness;

            if (!TryParseSvgHue(
                    components[0],
                    out hue) ||
                !TryParseSvgPercentComponent(
                    components[1],
                    out saturation) ||
                !TryParseSvgPercentComponent(
                    components[2],
                    out lightness))
            {
                return false;
            }

            int alpha =
                255;

            if (!string.IsNullOrEmpty(
                    alphaToken) &&
                !TryParseSvgAlphaChannel(
                    alphaToken,
                    out alpha))
            {
                return false;
            }

            float chroma =
                (1f -
                 Math.Abs(
                     2f *
                     lightness -
                     1f)) *
                saturation;
            float hueSection =
                hue /
                60f;
            float x =
                chroma *
                (1f -
                 Math.Abs(
                     hueSection %
                     2f -
                     1f));

            float r1 = 0f;
            float g1 = 0f;
            float b1 = 0f;

            if (hueSection < 1f)
            {
                r1 = chroma;
                g1 = x;
            }
            else if (hueSection < 2f)
            {
                r1 = x;
                g1 = chroma;
            }
            else if (hueSection < 3f)
            {
                g1 = chroma;
                b1 = x;
            }
            else if (hueSection < 4f)
            {
                g1 = x;
                b1 = chroma;
            }
            else if (hueSection < 5f)
            {
                r1 = x;
                b1 = chroma;
            }
            else
            {
                r1 = chroma;
                b1 = x;
            }

            float match =
                lightness -
                chroma /
                2f;

            int red =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            (r1 + match) *
                            255f)));
            int green =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            (g1 + match) *
                            255f)));
            int blue =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            (b1 + match) *
                            255f)));

            color =
                Color.FromArgb(
                    alpha,
                    red,
                    green,
                    blue);

            return true;
        }

        private static bool TryParseSvgHue(
            string raw,
            out float degrees)
        {
            degrees = 0f;

            if (string.IsNullOrEmpty(raw))
                return false;

            raw =
                raw.Trim();

            if (string.Equals(
                    raw,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            float multiplier =
                1f;

            if (raw.EndsWith(
                    "turn",
                    StringComparison.OrdinalIgnoreCase))
            {
                multiplier = 360f;
                raw =
                    raw.Substring(
                        0,
                        raw.Length - 4);
            }
            else if (raw.EndsWith(
                    "grad",
                    StringComparison.OrdinalIgnoreCase))
            {
                multiplier = 0.9f;
                raw =
                    raw.Substring(
                        0,
                        raw.Length - 4);
            }
            else if (raw.EndsWith(
                    "rad",
                    StringComparison.OrdinalIgnoreCase))
            {
                multiplier =
                    180f /
                    (float)Math.PI;
                raw =
                    raw.Substring(
                        0,
                        raw.Length - 3);
            }
            else if (raw.EndsWith(
                    "deg",
                    StringComparison.OrdinalIgnoreCase))
            {
                raw =
                    raw.Substring(
                        0,
                        raw.Length - 3);
            }

            float parsed =
                ParseSvgFloat(
                    raw,
                    float.NaN);

            if (float.IsNaN(parsed) ||
                float.IsInfinity(parsed))
            {
                return false;
            }

            degrees =
                parsed *
                multiplier;

            degrees =
                degrees %
                360f;

            if (degrees < 0f)
                degrees += 360f;

            return true;
        }

        private static bool TryParseSvgPercentComponent(
            string raw,
            out float value)
        {
            value = 0f;

            if (string.IsNullOrEmpty(raw))
                return false;

            raw =
                raw.Trim();

            if (string.Equals(
                    raw,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            bool percent =
                raw.EndsWith(
                    "%",
                    StringComparison.Ordinal);

            string number =
                percent
                    ? raw.Substring(
                        0,
                        raw.Length - 1)
                    : raw;

            float parsed =
                ParseSvgFloat(
                    number,
                    float.NaN);

            if (float.IsNaN(parsed) ||
                float.IsInfinity(parsed))
            {
                return false;
            }

            value =
                Math.Max(
                    0f,
                    Math.Min(
                        1f,
                        parsed /
                        100f));

            return true;
        }

        private static bool TryParseSvgAlphaChannel(
            string raw,
            out int value)
        {
            value = 255;

            if (string.IsNullOrEmpty(raw))
            {
                return false;
            }

            raw =
                raw.Trim();

            if (string.Equals(
                    raw,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                value = 0;
                return true;
            }

            bool percent =
                raw.EndsWith(
                    "%",
                    StringComparison.Ordinal);

            string number =
                percent
                    ? raw.Substring(
                        0,
                        raw.Length - 1)
                    : raw;

            float parsed =
                ParseSvgFloat(
                    number,
                    float.NaN);

            if (float.IsNaN(parsed) ||
                float.IsInfinity(parsed))
            {
                return false;
            }

            float alpha =
                percent
                    ? parsed /
                      100f
                    : parsed;

            value =
                Math.Max(
                    0,
                    Math.Min(
                        255,
                        (int)Math.Round(
                            Math.Max(
                                0f,
                                Math.Min(
                                    1f,
                                    alpha)) *
                            255f)));

            return true;
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
                    item,
                    theme);

                ReadChartSeriesLabelOverrides(
                    nodes[i],
                    item,
                    theme);

                if (item.Values.Count > 0)
                    series.Add(item);
            }

            return series;
        }

        private static void PrepareChartSurface(
            Graphics g,
            RectangleF rect,
            XmlDocument chartDoc,
            Dictionary<string, Color> theme,
            out RectangleF plot)
        {
            Color chartBackground =
                ReadChartAreaFill(
                    chartDoc,
                    theme,
                    "chartSpace",
                    Color.White);

            using (Brush bg =
                new SolidBrush(
                    chartBackground))
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

            bool titleOverlay =
                ReadChartOverlay(
                    chartDoc,
                    "title");

            if (!string.IsNullOrEmpty(
                    title))
            {
                ChartLabelOptions titleTextStyle =
                    ReadChartTitleTextStyle(
                        chartDoc,
                        theme);

                RectangleF titleRect =
                    new RectangleF(
                        rect.Left + 5f,
                        rect.Top + 4f,
                        rect.Width - 10f,
                        titleHeight - 4f);

                RectangleF manualTitleRect;

                if (TryResolveChartElementManualLayout(
                        chartDoc,
                        "title",
                        rect,
                        titleRect,
                        out manualTitleRect))
                {
                    titleRect =
                        manualTitleRect;
                }

                using (Font font =
                    SafeChartTextFont(
                        "Arial",
                        Math.Max(
                            10f,
                            Math.Min(
                                18f,
                                rect.Height /
                                18f)),
                        FontStyle.Bold,
                        titleTextStyle))
                using (Brush brush =
                    new SolidBrush(
                        titleTextStyle != null &&
                        titleTextStyle.TextColor.HasValue
                            ? titleTextStyle.TextColor.Value
                            : Color.FromArgb(
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
                        titleRect,
                        sf);
                }
            }

            string legendPosition =
                ReadChartLegendPosition(
                    chartDoc);
            bool legendOverlay =
                ReadChartOverlay(
                    chartDoc,
                    "legend");

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
                string.IsNullOrEmpty(
                    title)
                    ? 12f
                    : titleOverlay
                        ? 12f
                        : titleHeight;

            if (!legendOverlay)
            {
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

            ApplyChartManualPlotLayout(
                chartDoc,
                rect,
                ref plot);

            Color plotBackground =
                ReadChartAreaFill(
                    chartDoc,
                    theme,
                    "plotArea",
                    chartBackground);

            using (Brush plotBrush =
                new SolidBrush(
                    plotBackground))
            {
                g.FillRectangle(
                    plotBrush,
                    plot);
            }
        }

        private static void ApplyChartManualPlotLayout(
            XmlDocument chartDoc,
            RectangleF chartRect,
            ref RectangleF plot)
        {
            if (chartDoc == null)
                return;

            XmlNode plotArea =
                FindFirst(
                    chartDoc,
                    "plotArea");

            XmlNode layout =
                plotArea == null
                    ? null
                    : DirectChild(
                        plotArea,
                        "layout");

            XmlNode manualLayout =
                layout == null
                    ? null
                    : DirectChild(
                        layout,
                        "manualLayout");

            if (manualLayout == null)
                return;

            RectangleF automaticInner =
                plot;

            string layoutTarget =
                ReadChartManualLayoutTarget(
                    manualLayout);

            if (layoutTarget == "inner" ||
                !ChartHasRenderableAxes(
                    chartDoc))
            {
                plot =
                    ResolveChartManualLayoutRectangle(
                        manualLayout,
                        chartRect,
                        automaticInner);

                return;
            }

            RectangleF automaticOuter =
                EstimateChartAutomaticOuterPlotRectangle(
                    chartDoc,
                    chartRect,
                    automaticInner);

            RectangleF resolvedOuter =
                ResolveChartManualLayoutRectangle(
                    manualLayout,
                    chartRect,
                    automaticOuter);

            plot =
                ConvertChartOuterLayoutToInner(
                    chartRect,
                    automaticOuter,
                    automaticInner,
                    resolvedOuter);
        }

        private static string ReadChartManualLayoutTarget(
            XmlNode manualLayout)
        {
            if (manualLayout == null)
                return "outer";

            XmlNode target =
                DirectChild(
                    manualLayout,
                    "layoutTarget");

            string value =
                target == null
                    ? string.Empty
                    : GetAttr(
                        target,
                        "val");

            return string.Equals(
                    value,
                    "inner",
                    StringComparison.OrdinalIgnoreCase)
                ? "inner"
                : "outer";
        }

        private static bool ChartHasRenderableAxes(
            XmlDocument chartDoc)
        {
            if (chartDoc == null)
                return false;

            XmlNode plotArea =
                FindFirst(
                    chartDoc,
                    "plotArea");

            if (plotArea == null)
                return false;

            return
                FindFirst(
                    plotArea,
                    "catAx") != null ||
                FindFirst(
                    plotArea,
                    "valAx") != null ||
                FindFirst(
                    plotArea,
                    "dateAx") != null ||
                FindFirst(
                    plotArea,
                    "serAx") != null;
        }

        private static RectangleF EstimateChartAutomaticOuterPlotRectangle(
            XmlDocument chartDoc,
            RectangleF chartRect,
            RectangleF automaticInner)
        {
            float left =
                chartRect.Left +
                Math.Max(
                    6f,
                    chartRect.Width *
                    0.015f);

            float right =
                chartRect.Right -
                Math.Max(
                    6f,
                    chartRect.Width *
                    0.015f);

            float top =
                chartRect.Top +
                12f;

            float bottom =
                chartRect.Bottom -
                Math.Max(
                    8f,
                    chartRect.Height *
                    0.025f);

            string title =
                ReadChartTitle(
                    chartDoc);

            if (!string.IsNullOrEmpty(
                    title) &&
                !ReadChartOverlay(
                    chartDoc,
                    "title"))
            {
                top =
                    Math.Max(
                        top,
                        chartRect.Top +
                        Math.Max(
                            34f,
                            chartRect.Height *
                            0.10f));
            }

            if (!ReadChartOverlay(
                    chartDoc,
                    "legend"))
            {
                string legendPosition =
                    ReadChartLegendPosition(
                        chartDoc);

                if (legendPosition == "l")
                {
                    left =
                        Math.Max(
                            left,
                            chartRect.Left +
                            Math.Max(
                                82f,
                                chartRect.Width *
                                0.19f));
                }
                else if (legendPosition == "r" ||
                         legendPosition == "tr")
                {
                    right =
                        Math.Min(
                            right,
                            chartRect.Right -
                            Math.Max(
                                82f,
                                chartRect.Width *
                                0.19f));
                }
                else if (legendPosition == "t")
                {
                    top +=
                        Math.Max(
                            28f,
                            chartRect.Height *
                            0.09f);
                }
                else if (legendPosition == "b")
                {
                    bottom -=
                        Math.Max(
                            58f,
                            chartRect.Height *
                            0.17f);
                }
            }

            if (right <= left + 12f ||
                bottom <= top + 12f)
            {
                return automaticInner;
            }

            RectangleF automaticOuter =
                new RectangleF(
                    left,
                    top,
                    right - left,
                    bottom - top);

            if (automaticInner.Left <
                    automaticOuter.Left ||
                automaticInner.Top <
                    automaticOuter.Top ||
                automaticInner.Right >
                    automaticOuter.Right ||
                automaticInner.Bottom >
                    automaticOuter.Bottom)
            {
                return automaticInner;
            }

            return automaticOuter;
        }

        private static RectangleF ConvertChartOuterLayoutToInner(
            RectangleF chartRect,
            RectangleF automaticOuter,
            RectangleF automaticInner,
            RectangleF resolvedOuter)
        {
            if (automaticOuter.Width <= 0f ||
                automaticOuter.Height <= 0f)
            {
                return automaticInner;
            }

            float leftRatio =
                Math.Max(
                    0f,
                    (automaticInner.Left -
                     automaticOuter.Left) /
                    automaticOuter.Width);

            float topRatio =
                Math.Max(
                    0f,
                    (automaticInner.Top -
                     automaticOuter.Top) /
                    automaticOuter.Height);

            float rightRatio =
                Math.Max(
                    0f,
                    (automaticOuter.Right -
                     automaticInner.Right) /
                    automaticOuter.Width);

            float bottomRatio =
                Math.Max(
                    0f,
                    (automaticOuter.Bottom -
                     automaticInner.Bottom) /
                    automaticOuter.Height);

            float left =
                resolvedOuter.Left +
                resolvedOuter.Width *
                leftRatio;

            float top =
                resolvedOuter.Top +
                resolvedOuter.Height *
                topRatio;

            float right =
                resolvedOuter.Right -
                resolvedOuter.Width *
                rightRatio;

            float bottom =
                resolvedOuter.Bottom -
                resolvedOuter.Height *
                bottomRatio;

            left =
                Math.Max(
                    chartRect.Left,
                    Math.Min(
                        chartRect.Right - 12f,
                        left));

            top =
                Math.Max(
                    chartRect.Top,
                    Math.Min(
                        chartRect.Bottom - 12f,
                        top));

            right =
                Math.Max(
                    left + 12f,
                    Math.Min(
                        chartRect.Right,
                        right));

            bottom =
                Math.Max(
                    top + 12f,
                    Math.Min(
                        chartRect.Bottom,
                        bottom));

            return new RectangleF(
                left,
                top,
                right - left,
                bottom - top);
        }

        private static bool TryResolveChartElementManualLayout(
            XmlDocument chartDoc,
            string elementName,
            RectangleF chartRect,
            RectangleF automaticRect,
            out RectangleF resolvedRect)
        {
            resolvedRect =
                automaticRect;

            if (chartDoc == null ||
                string.IsNullOrEmpty(
                    elementName))
            {
                return false;
            }

            XmlNode chart =
                FindFirst(
                    chartDoc,
                    "chart");

            XmlNode element =
                chart == null
                    ? null
                    : DirectChild(
                        chart,
                        elementName);

            XmlNode layout =
                element == null
                    ? null
                    : DirectChild(
                        element,
                        "layout");

            XmlNode manualLayout =
                layout == null
                    ? null
                    : DirectChild(
                        layout,
                        "manualLayout");

            if (manualLayout == null)
                return false;

            resolvedRect =
                ResolveChartManualLayoutRectangle(
                    manualLayout,
                    chartRect,
                    automaticRect);

            return true;
        }

        private static RectangleF ResolveChartManualLayoutRectangle(
            XmlNode manualLayout,
            RectangleF chartRect,
            RectangleF automaticRect)
        {
            if (manualLayout == null ||
                !IsChartManualLayoutRangeValid(
                    manualLayout))
            {
                return automaticRect;
            }

            double x =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "x");

            double y =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "y");

            double width =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "w");

            double height =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "h");

            string xMode =
                ReadChartManualLayoutMode(
                    manualLayout,
                    "xMode");

            string yMode =
                ReadChartManualLayoutMode(
                    manualLayout,
                    "yMode");

            string widthMode =
                ReadChartManualLayoutMode(
                    manualLayout,
                    "wMode");

            string heightMode =
                ReadChartManualLayoutMode(
                    manualLayout,
                    "hMode");

            float left =
                automaticRect.Left;

            float top =
                automaticRect.Top;

            if (!double.IsNaN(x))
            {
                left =
                    xMode == "edge"
                        ? chartRect.Left +
                            (float)x *
                            chartRect.Width
                        : automaticRect.Left +
                            (float)x *
                            chartRect.Width;
            }

            if (!double.IsNaN(y))
            {
                top =
                    yMode == "edge"
                        ? chartRect.Top +
                            (float)y *
                            chartRect.Height
                        : automaticRect.Top +
                            (float)y *
                            chartRect.Height;
            }

            float right =
                left +
                automaticRect.Width;

            float bottom =
                top +
                automaticRect.Height;

            if (!double.IsNaN(width))
            {
                right =
                    widthMode == "edge"
                        ? chartRect.Left +
                            (float)width *
                            chartRect.Width
                        : left +
                            (float)width *
                            chartRect.Width;
            }

            if (!double.IsNaN(height))
            {
                bottom =
                    heightMode == "edge"
                        ? chartRect.Top +
                            (float)height *
                            chartRect.Height
                        : top +
                            (float)height *
                            chartRect.Height;
            }

            float resolvedWidth =
                right -
                left;

            float resolvedHeight =
                bottom -
                top;

            if (resolvedWidth <= 0f ||
                resolvedHeight <= 0f)
            {
                return automaticRect;
            }

            resolvedWidth =
                Math.Max(
                    12f,
                    Math.Min(
                        chartRect.Width,
                        resolvedWidth));

            resolvedHeight =
                Math.Max(
                    12f,
                    Math.Min(
                        chartRect.Height,
                        resolvedHeight));

            if (left <
                chartRect.Left)
            {
                left =
                    chartRect.Left;
            }

            if (top <
                chartRect.Top)
            {
                top =
                    chartRect.Top;
            }

            if (left + resolvedWidth >
                chartRect.Right)
            {
                left =
                    chartRect.Right -
                    resolvedWidth;
            }

            if (top + resolvedHeight >
                chartRect.Bottom)
            {
                top =
                    chartRect.Bottom -
                    resolvedHeight;
            }

            left =
                Math.Max(
                    chartRect.Left,
                    left);

            top =
                Math.Max(
                    chartRect.Top,
                    top);

            return new RectangleF(
                left,
                top,
                resolvedWidth,
                resolvedHeight);
        }

        private static bool IsChartManualLayoutRangeValid(
            XmlNode manualLayout)
        {
            if (manualLayout == null)
                return false;

            double x =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "x");

            double y =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "y");

            double width =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "w");

            double height =
                ReadChartManualLayoutNumber(
                    manualLayout,
                    "h");

            string xMode =
                ReadChartManualLayoutMode(
                    manualLayout,
                    "xMode");

            string yMode =
                ReadChartManualLayoutMode(
                    manualLayout,
                    "yMode");

            if (!double.IsNaN(x) &&
                (x > 1.0 ||
                 (xMode == "edge" &&
                  x < 0.0) ||
                 (xMode == "factor" &&
                  x < -1.0)))
            {
                return false;
            }

            if (!double.IsNaN(y) &&
                (y > 1.0 ||
                 (yMode == "edge" &&
                  y < 0.0) ||
                 (yMode == "factor" &&
                  y < -1.0)))
            {
                return false;
            }

            if (!double.IsNaN(width) &&
                (width < 0.0 ||
                 width > 1.0))
            {
                return false;
            }

            if (!double.IsNaN(height) &&
                (height < 0.0 ||
                 height > 1.0))
            {
                return false;
            }

            return true;
        }

        private static double ReadChartManualLayoutNumber(
            XmlNode manualLayout,
            string localName)
        {
            if (manualLayout == null)
                return double.NaN;

            XmlNode node =
                FindFirst(
                    manualLayout,
                    localName);

            string raw =
                node == null
                    ? string.Empty
                    : GetAttr(
                        node,
                        "val");

            double value;

            if (!double.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out value))
            {
                return double.NaN;
            }

            return value;
        }

        private static string ReadChartManualLayoutMode(
            XmlNode manualLayout,
            string localName)
        {
            if (manualLayout == null)
                return "factor";

            XmlNode node =
                FindFirst(
                    manualLayout,
                    localName);

            string value =
                node == null
                    ? string.Empty
                    : GetAttr(
                        node,
                        "val");

            return string.Equals(
                    value,
                    "edge",
                    StringComparison.OrdinalIgnoreCase)
                ? "edge"
                : "factor";
        }

        private static void DrawDoughnutChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            RectangleF plot;
            PrepareChartSurface(g, rect, chartDoc, theme, out plot);
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
            Color doughnutHoleColor =
                ReadChartAreaFill(
                    chartDoc,
                    theme,
                    "plotArea",
                    ReadChartAreaFill(
                        chartDoc,
                        theme,
                        "chartSpace",
                        Color.White));

            using (Brush holeBrush =
                new SolidBrush(
                    doughnutHoleColor))
            {
                g.FillEllipse(
                    holeBrush,
                    hole);
            }

            ChartLabelOptions labels =
                ReadChartLabelOptions(
                    chartDoc,
                    theme);
            labels.LayoutReferenceRect =
                rect;

            if (HasAnyChartSeriesDataLabel(
                    labels,
                    data))
            {
                DrawDoughnutChartValueLabels(
                    g,
                    pie,
                    hole,
                    data,
                    labels,
                    ReadChartFirstSliceAngle(
                        chartDoc),
                    ReadChartLeaderLineStyle(
                        chartDoc,
                        theme));
            }

            DrawChartLegend(
                g,
                rect,
                series,
                palette,
                "pie",
                ReadChartLegendPosition(chartDoc),
                ReadChartLegendTextStyle(
                    chartDoc,
                    theme),
                ReadChartLegendHiddenEntries(
                    chartDoc),
                    chartDoc);
        }

        private static void DrawDoughnutChartValueLabels(
            Graphics g,
            RectangleF outer,
            RectangleF hole,
            ChartSeriesData series,
            ChartLabelOptions options,
            float startAngle,
            ChartLineStyle leaderLineStyle)
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

                    ChartLabelOptions pointLabels =
                        ResolveChartPointLabelOptions(
                            options,
                            series,
                            i);

                    if (pointLabels.HasAny)
                    {
                        string normalizedPosition =
                            NormalizeChartDataLabelPosition(
                                pointLabels.Position);

                        float radialFactor =
                            normalizedPosition == "outend"
                                ? 1.10f
                                : normalizedPosition == "ctr"
                                    ? 0.50f
                                    : normalizedPosition == "inend"
                                        ? 0.72f
                                        : 0.55f;

                        float labelRadius =
                            normalizedPosition == "outend"
                                ? outerRadius *
                                  radialFactor
                                : innerRadius +
                                  (outerRadius -
                                   innerRadius) *
                                  radialFactor;

                        string label =
                            BuildChartDataLabel(
                                pointLabels,
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

                            if (normalizedPosition == "outend" &&
                                pointLabels.ShowLeaderLines)
                            {
                                if (leaderLineStyle == null)
                                {
                                    leaderLineStyle =
                                        new ChartLineStyle();
                                    leaderLineStyle.Color =
                                        Color.FromArgb(
                                            120,
                                            120,
                                            120);
                                    leaderLineStyle.Width = 1f;
                                }

                                PointF lineStart =
                                    new PointF(
                                        cx +
                                        (float)Math.Cos(
                                            radians) *
                                        outerRadius *
                                        0.96f,
                                        cy +
                                        (float)Math.Sin(
                                            radians) *
                                        outerRadius *
                                        0.96f);
                                PointF lineEnd =
                                    new PointF(
                                        cx +
                                        (float)Math.Cos(
                                            radians) *
                                        outerRadius *
                                        1.05f,
                                        cy +
                                        (float)Math.Sin(
                                            radians) *
                                        outerRadius *
                                        1.05f);

                                using (Pen leader =
                                    new Pen(
                                        leaderLineStyle.Color,
                                        leaderLineStyle.Width))
                                {
                                    leader.DashStyle =
                                        leaderLineStyle.DashStyle;

                                    g.DrawLine(
                                        leader,
                                        lineStart,
                                        lineEnd);
                                }
                            }

                            DrawChartLabelText(
                                g,
                                label,
                                font,
                                brush,
                                x,
                                y,
                                pointLabels);
                        }
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
                theme,
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

            bool valueAxisDeleted =
                ReadChartAxisDeleted(
                    chartDoc,
                    "valAx");
            bool categoryAxisDeleted =
                ReadChartAxisDeleted(
                    chartDoc,
                    "catAx");

            string valueTickLabelPosition =
                ResolveChartAxisTickLabelPosition(
                    chartDoc,
                    "valAx",
                    "column");

            DrawChartValueGrid(
                g,
                plot,
                axisScale,
                "column",
                !valueAxisDeleted &&
                !string.Equals(
                    valueTickLabelPosition,
                    "none",
                    StringComparison.OrdinalIgnoreCase),
                valueTickLabelPosition,
                ReadChartMajorGridlineStyle(
                    chartDoc,
                    theme),
                ReadChartMinorGridlineStyle(
                    chartDoc,
                    theme),
                ReadChartAxisLabelStyle(
                    chartDoc,
                    "valAx",
                    theme));

            double categoryAxisCrossValue =
                ReadChartCategoryAxisCrossValue(
                    chartDoc,
                    axisScale);

            float zeroY =
                plot.Bottom -
                (float)(
                    ChartAxisFraction(
                        categoryAxisCrossValue,
                        axisScale) *
                    plot.Height);

            zeroY =
                Math.Max(
                    plot.Top,
                    Math.Min(
                        plot.Bottom,
                        zeroY));

            ChartLineStyle categoryAxisStyle =
                ReadChartAxisLineStyle(
                    chartDoc,
                    "catAx",
                    theme);
            ChartLineStyle valueAxisStyle =
                ReadChartAxisLineStyle(
                    chartDoc,
                    "valAx",
                    theme);

            using (Pen categoryAxis =
                new Pen(
                    categoryAxisStyle.Color,
                    categoryAxisStyle.Width))
            using (Pen valueAxis =
                new Pen(
                    valueAxisStyle.Color,
                    valueAxisStyle.Width))
            {
                categoryAxis.DashStyle =
                    categoryAxisStyle.DashStyle;
                valueAxis.DashStyle =
                    valueAxisStyle.DashStyle;

                if (!categoryAxisDeleted)
                {
                    g.DrawLine(
                        categoryAxis,
                        plot.Left,
                        zeroY,
                        plot.Right,
                        zeroY);
                }

                if (!valueAxisDeleted)
                {
                    string valueAxisPosition =
                        ReadChartAxisPosition(
                            chartDoc,
                            "valAx");

                    float valueAxisX =
                        valueAxisPosition == "r"
                            ? plot.Right
                            : plot.Left;

                    g.DrawLine(
                        valueAxis,
                        valueAxisX,
                        plot.Top,
                        valueAxisX,
                        plot.Bottom);
                }
            }

            DrawChartAxisTickMarks(
                g,
                plot,
                axisScale,
                categoryCount,
                "line",
                0f,
                zeroY,
                chartDoc,
                categoryAxisStyle,
                valueAxisStyle);

            Color[] palette =
                EnhancedChartPalette(
                    theme);

            ChartLabelOptions labels =
                ReadChartLabelOptions(
                    chartDoc,
                    theme);
            labels.LayoutReferenceRect =
                rect;

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

                    if (!stacked)
                    {
                        DrawChartTrendline(
                            g,
                            plot,
                            topPoints,
                            item.Trendline,
                            color);

                        for (int i = 0;
                             i < topPoints.Count &&
                             i < item.Values.Count;
                             i++)
                        {
                            DrawChartErrorBar(
                                g,
                                plot,
                                axisScale,
                                "line",
                                topPoints[i],
                                item.Values[i],
                                i,
                                item.ErrorBars,
                                color);
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

                    if (HasAnyChartSeriesDataLabel(
                            labels,
                            item))
                    {
                        for (int i = 0;
                             i < item.Values.Count &&
                             i < topPoints.Count;
                             i++)
                        {
                            ChartLabelOptions pointLabels =
                                ResolveChartPointLabelOptions(
                                    labels,
                                    item,
                                    i);

                            if (!pointLabels.HasAny)
                                continue;

                            string label =
                                BuildChartDataLabel(
                                    pointLabels,
                                    item,
                                    i,
                                    false);

                            DrawChartPointDataLabel(
                                g,
                                labelFont,
                                labelBrush,
                                label,
                                topPoints[i],
                                pointLabels);
                        }
                    }
                }
            }

            string categoryTickLabelPosition =
                ResolveChartAxisTickLabelPosition(
                    chartDoc,
                    "catAx",
                    "column");

            if (!categoryAxisDeleted &&
                !string.Equals(
                    categoryTickLabelPosition,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                DrawChartCategoryLabels(
                    g,
                    plot,
                    series[0],
                    categoryCount,
                    "column",
                    categoryTickLabelPosition,
                    ReadChartCategoryLabelSkip(
                        chartDoc),
                    ReadChartAxisLabelStyle(
                        chartDoc,
                        "catAx",
                        theme));
            }

            DrawChartAxisTitles(
                g,
                plot,
                rect,
                "column",
                valueAxisDeleted
                    ? string.Empty
                    : ReadChartAxisTitle(
                        chartDoc,
                        "valAx"),
                categoryAxisDeleted
                    ? string.Empty
                    : ReadChartAxisTitle(
                        chartDoc,
                        "catAx"),
                ReadChartAxisTitleTextStyle(
                    chartDoc,
                    "valAx",
                    theme),
                ReadChartAxisTitleTextStyle(
                    chartDoc,
                    "catAx",
                    theme));

            DrawChartLegend(
                g,
                rect,
                series,
                palette,
                "area",
                ReadChartLegendPosition(
                    chartDoc),
                ReadChartLegendTextStyle(
                    chartDoc,
                    theme),
                ReadChartLegendHiddenEntries(
                    chartDoc),
                    chartDoc);
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
                theme,
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
                    style,
                    theme);

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

            ChartAxisScale scatterYScale =
                new ChartAxisScale();
            scatterYScale.Minimum =
                minY;
            scatterYScale.Maximum =
                maxY;
            scatterYScale.MajorUnit = 0.0;
            scatterYScale.MinorUnit = 0.0;

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

                DrawChartTrendline(
                    g,
                    plot,
                    points,
                    style.Trendline,
                    seriesColor);

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

                    if (!string.Equals(
                            style.ErrorBars == null
                                ? string.Empty
                                : style.ErrorBars.Direction,
                            "x",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        DrawChartErrorBar(
                            g,
                            plot,
                            scatterYScale,
                            "line",
                            points[i],
                            allY[s][i],
                            i,
                            style.ErrorBars,
                            pointColor);
                    }

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
                        chartDoc),
                    ReadChartLegendTextStyle(
                        chartDoc,
                        theme),
                    ReadChartLegendHiddenEntries(
                        chartDoc),
                        chartDoc);
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
            PrepareChartSurface(g, rect, chartDoc, theme, out plot);
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
                ReadChartLegendPosition(chartDoc),
                ReadChartLegendTextStyle(
                    chartDoc,
                    theme),
                ReadChartLegendHiddenEntries(
                    chartDoc),
                    chartDoc);
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

                if (type == "doc")
                {
                    continue;
                }

                SmartNode node =
                    new SmartNode();
                node.Id = id;
                node.IsAssistant =
                    string.Equals(
                        type,
                        "asst",
                        StringComparison.OrdinalIgnoreCase);
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

                SmartNode destinationNode;
                if (nodes.TryGetValue(
                        destination,
                        out destinationNode) &&
                    string.IsNullOrEmpty(
                        destinationNode.ParentId))
                {
                    destinationNode.ParentId =
                        source;
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
                if (!orderedNodes[i].IsAssistant)
                {
                    maxDepth =
                        Math.Max(
                            maxDepth,
                            orderedNodes[i].Depth);
                }
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

                if (node.IsAssistant)
                {
                    accent =
                        LightenSmartArtColor(
                            accent,
                            0.18f);
                }

                Color fill =
                    LightenSmartArtColor(
                        accent,
                        node.IsAssistant
                            ? 0.90f
                            : 0.82f);

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
                    if (node.IsAssistant)
                    {
                        border.DashStyle =
                            DashStyle.Dash;
                    }

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
                    node.IsAssistant
                        ? FontStyle.Italic
                        : node.Depth == 0
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
                    if (nodes[i].Depth == depth &&
                        !nodes[i].IsAssistant)
                    {
                        level.Add(nodes[i]);
                    }
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

                float nodeGap =
                    Math.Max(
                        8f,
                        Math.Min(
                            rect.Width,
                            rect.Height) *
                        0.018f);

                Dictionary<string, List<SmartNode>> groups =
                    new Dictionary<string, List<SmartNode>>(
                        StringComparer.Ordinal);
                Dictionary<string, float> groupCenters =
                    new Dictionary<string, float>(
                        StringComparer.Ordinal);
                List<string> groupOrder =
                    new List<string>();

                for (int i = 0;
                     i < level.Count;
                     i++)
                {
                    SmartNode item =
                        level[i];

                    RectangleF parentBox =
                        RectangleF.Empty;
                    bool hasParent =
                        depth > 0 &&
                        !string.IsNullOrEmpty(
                            item.ParentId) &&
                        positions.TryGetValue(
                            item.ParentId,
                            out parentBox);

                    string groupKey =
                        hasParent
                            ? item.ParentId
                            : "__level_" +
                              depth.ToString() +
                              "_" +
                              i.ToString();

                    List<SmartNode> siblings;
                    if (!groups.TryGetValue(
                            groupKey,
                            out siblings))
                    {
                        siblings =
                            new List<SmartNode>();
                        groups[groupKey] =
                            siblings;
                        groupOrder.Add(
                            groupKey);

                        groupCenters[groupKey] =
                            hasParent
                                ? parentBox.Left +
                                  parentBox.Width /
                                  2f
                                : rect.Left +
                                  (i + 0.5f) *
                                  columnWidth;
                    }

                    siblings.Add(
                        item);
                }

                groupOrder.Sort(
                    delegate(
                        string left,
                        string right)
                    {
                        return groupCenters[left]
                            .CompareTo(
                                groupCenters[right]);
                    });

                float y =
                    rect.Top +
                    depth * rowHeight +
                    (rowHeight -
                     boxH) /
                    2f;
                float previousRight =
                    rect.Left -
                    nodeGap;
                List<string> placedIds =
                    new List<string>();

                for (int groupIndex = 0;
                     groupIndex < groupOrder.Count;
                     groupIndex++)
                {
                    string groupKey =
                        groupOrder[groupIndex];
                    List<SmartNode> siblings =
                        groups[groupKey];

                    float groupWidth =
                        siblings.Count *
                        boxW +
                        Math.Max(
                            0,
                            siblings.Count - 1) *
                        nodeGap;

                    float x =
                        groupCenters[groupKey] -
                        groupWidth /
                        2f;

                    x =
                        Math.Max(
                            rect.Left,
                            Math.Max(
                                previousRight +
                                nodeGap,
                                x));

                    for (int siblingIndex = 0;
                         siblingIndex < siblings.Count;
                         siblingIndex++)
                    {
                        float childX =
                            x +
                            siblingIndex *
                            (boxW +
                             nodeGap);

                        positions[
                            siblings[siblingIndex].Id] =
                            new RectangleF(
                                childX,
                                y,
                                boxW,
                                boxH);

                        placedIds.Add(
                            siblings[siblingIndex].Id);
                    }

                    previousRight =
                        x +
                        groupWidth;
                }

                if (previousRight >
                    rect.Right &&
                    placedIds.Count > 0)
                {
                    float shift =
                        previousRight -
                        rect.Right;

                    for (int i = 0;
                         i < placedIds.Count;
                         i++)
                    {
                        RectangleF box =
                            positions[
                                placedIds[i]];
                        box.X -= shift;
                        positions[
                            placedIds[i]] =
                            box;
                    }

                    float firstLeft =
                        positions[
                            placedIds[0]]
                            .Left;

                    if (firstLeft <
                        rect.Left)
                    {
                        float correction =
                            rect.Left -
                            firstLeft;

                        for (int i = 0;
                             i < placedIds.Count;
                             i++)
                        {
                            RectangleF box =
                                positions[
                                    placedIds[i]];
                            box.X +=
                                correction;
                            positions[
                                placedIds[i]] =
                                box;
                        }
                    }
                }
            }

            Dictionary<string, int> assistantCounts =
                new Dictionary<string, int>(
                    StringComparer.Ordinal);
            Dictionary<string, int> assistantTotals =
                new Dictionary<string, int>(
                    StringComparer.Ordinal);

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                SmartNode candidate =
                    nodes[i];

                if (!candidate.IsAssistant ||
                    string.IsNullOrEmpty(
                        candidate.ParentId))
                {
                    continue;
                }

                int total = 0;
                assistantTotals.TryGetValue(
                    candidate.ParentId,
                    out total);
                assistantTotals[
                    candidate.ParentId] =
                    total + 1;
            }

            for (int i = 0;
                 i < nodes.Count;
                 i++)
            {
                SmartNode assistant =
                    nodes[i];

                if (!assistant.IsAssistant)
                    continue;

                RectangleF parentBox;
                if (string.IsNullOrEmpty(
                        assistant.ParentId) ||
                    !positions.TryGetValue(
                        assistant.ParentId,
                        out parentBox))
                {
                    continue;
                }

                int assistantIndex = 0;
                assistantCounts.TryGetValue(
                    assistant.ParentId,
                    out assistantIndex);
                assistantCounts[
                    assistant.ParentId] =
                    assistantIndex + 1;

                float gap =
                    Math.Max(
                        8f,
                        Math.Min(
                            rect.Width,
                            rect.Height) *
                        0.025f);

                float leftSpace =
                    parentBox.Left -
                    rect.Left;
                float rightSpace =
                    rect.Right -
                    parentBox.Right;

                float desiredBoxW =
                    Math.Max(
                        42f,
                        parentBox.Width *
                        0.62f);
                float sideCapacity =
                    Math.Max(
                        0f,
                        Math.Max(
                            leftSpace,
                            rightSpace) -
                        gap);

                float boxW =
                    sideCapacity >= 42f
                        ? Math.Min(
                            desiredBoxW,
                            sideCapacity)
                        : Math.Max(
                            32f,
                            Math.Min(
                                parentBox.Width *
                                    0.42f,
                                rect.Width *
                                    0.20f));

                float boxH =
                    Math.Max(
                        24f,
                        parentBox.Height *
                        0.72f);

                bool placeLeft =
                    assistantIndex % 2 == 0;

                if (placeLeft &&
                    leftSpace <
                        boxW + gap &&
                    rightSpace >
                        leftSpace)
                {
                    placeLeft = false;
                }
                else if (!placeLeft &&
                         rightSpace <
                            boxW + gap &&
                         leftSpace >
                            rightSpace)
                {
                    placeLeft = true;
                }

                float x =
                    placeLeft
                        ? parentBox.Left -
                          gap -
                          boxW
                        : parentBox.Right +
                          gap;

                x =
                    Math.Max(
                        rect.Left,
                        Math.Min(
                            rect.Right -
                            boxW,
                            x));

                int totalAssistants = 1;
                assistantTotals.TryGetValue(
                    assistant.ParentId,
                    out totalAssistants);

                int sideIndex =
                    assistantIndex /
                    2;
                int sideTotal =
                    placeLeft
                        ? (totalAssistants + 1) /
                          2
                        : totalAssistants /
                          2;

                sideTotal =
                    Math.Max(
                        1,
                        sideTotal);

                float stackGap =
                    Math.Max(
                        5f,
                        boxH *
                        0.18f);
                float stackHeight =
                    sideTotal *
                    boxH +
                    Math.Max(
                        0,
                        sideTotal - 1) *
                    stackGap;

                float y =
                    parentBox.Top +
                    parentBox.Height /
                    2f -
                    stackHeight /
                    2f +
                    sideIndex *
                    (boxH +
                     stackGap);

                y =
                    Math.Max(
                        rect.Top,
                        Math.Min(
                            rect.Bottom -
                            boxH,
                            y));

                positions[
                    assistant.Id] =
                    new RectangleF(
                        x,
                        y,
                        boxW,
                        boxH);
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

            float gap =
                Math.Max(
                    8f,
                    Math.Min(
                        rect.Width,
                        rect.Height) *
                    0.025f);

            float minimumCellWidth =
                Math.Max(
                    24f,
                    Math.Min(
                        40f,
                        rect.Width -
                        gap *
                        2f));

            int maxColumnsByWidth =
                Math.Max(
                    1,
                    (int)Math.Floor(
                        (rect.Width -
                         gap) /
                        Math.Max(
                            1f,
                            minimumCellWidth +
                            gap)));

            int columns =
                Math.Max(
                    1,
                    Math.Min(
                        Math.Min(
                            5,
                            count),
                        maxColumnsByWidth));

            int rows =
                (int)Math.Ceiling(
                    count /
                    (double)columns);

            float cellW =
                Math.Max(
                    minimumCellWidth,
                    (rect.Width -
                     gap * (columns + 1)) /
                    columns);

            cellW =
                Math.Min(
                    cellW,
                    Math.Max(
                        20f,
                        rect.Width -
                        gap *
                        2f));

            float cellH =
                Math.Max(
                    28f,
                    (rect.Height -
                     gap * (rows + 1)) /
                    rows);

            cellH =
                Math.Min(
                    cellH,
                    Math.Max(
                        44f,
                        rect.Height *
                        0.28f));

            float totalHeight =
                rows *
                cellH +
                Math.Max(
                    0,
                    rows - 1) *
                gap;
            float startY =
                rect.Top +
                Math.Max(
                    gap,
                    (rect.Height -
                     totalHeight) /
                    2f);

            for (int row = 0;
                 row < rows;
                 row++)
            {
                int rowStart =
                    row *
                    columns;
                int nodesInRow =
                    Math.Min(
                        columns,
                        nodes.Count -
                        rowStart);

                if (nodesInRow <= 0)
                    continue;

                float rowWidth =
                    nodesInRow *
                    cellW +
                    Math.Max(
                        0,
                        nodesInRow - 1) *
                    gap;
                float rowStartX =
                    rect.Left +
                    (rect.Width -
                     rowWidth) /
                    2f;

                for (int column = 0;
                     column < nodesInRow;
                     column++)
                {
                    int index =
                        rowStart +
                        column;

                    positions[
                        nodes[index].Id] =
                        new RectangleF(
                            rowStartX +
                            column *
                            (cellW + gap),
                            startY +
                            row *
                            (cellH + gap),
                            cellW,
                            cellH);
                }
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
                    20f,
                    Math.Min(
                        Math.Max(
                            60f,
                            rect.Width *
                            0.72f),
                        Math.Max(
                            20f,
                            rect.Width -
                            gap *
                            2f)));

            float totalHeight =
                count *
                boxH +
                Math.Max(
                    0,
                    count - 1) *
                gap;
            float startY =
                rect.Top +
                Math.Max(
                    gap,
                    (rect.Height -
                     totalHeight) /
                    2f);

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
                    startY +
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

                List<PointF[]> routedElbows =
                    new List<PointF[]>();

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
                        bool processSameRow =
                            false;

                        SmartNode destinationNode;
                        bool destinationIsAssistant =
                            nodes.TryGetValue(
                                node.Children[i],
                                out destinationNode) &&
                            destinationNode.IsAssistant;

                        if (layoutKind == "process")
                        {
                            float sourceCenterX =
                                source.Left +
                                source.Width /
                                2f;
                            float sourceCenterY =
                                source.Top +
                                source.Height /
                                2f;
                            float destinationCenterX =
                                destination.Left +
                                destination.Width /
                                2f;
                            float destinationCenterY =
                                destination.Top +
                                destination.Height /
                                2f;

                            processSameRow =
                                Math.Abs(
                                    sourceCenterY -
                                    destinationCenterY) <=
                                Math.Max(
                                    source.Height,
                                    destination.Height) *
                                0.45f;

                            if (processSameRow)
                            {
                                bool destinationOnRight =
                                    destinationCenterX >=
                                    sourceCenterX;

                                from =
                                    new PointF(
                                        destinationOnRight
                                            ? source.Right
                                            : source.Left,
                                        sourceCenterY);

                                to =
                                    new PointF(
                                        destinationOnRight
                                            ? destination.Left
                                            : destination.Right,
                                        destinationCenterY);
                            }
                            else
                            {
                                bool destinationBelow =
                                    destinationCenterY >=
                                    sourceCenterY;

                                from =
                                    new PointF(
                                        sourceCenterX,
                                        destinationBelow
                                            ? source.Bottom
                                            : source.Top);

                                to =
                                    new PointF(
                                        destinationCenterX,
                                        destinationBelow
                                            ? destination.Top
                                            : destination.Bottom);
                            }
                        }
                        else if (layoutKind == "hierarchy" &&
                                 destinationIsAssistant)
                        {
                            bool assistantOnLeft =
                                destination.Left <
                                source.Left;

                            from =
                                new PointF(
                                    assistantOnLeft
                                        ? source.Left
                                        : source.Right,
                                    source.Top +
                                    source.Height /
                                    2f);

                            to =
                                new PointF(
                                    assistantOnLeft
                                        ? destination.Right
                                        : destination.Left,
                                    destination.Top +
                                    destination.Height /
                                    2f);
                        }
                        else if (layoutKind == "hierarchy" ||
                                 layoutKind == "verticalProcess")
                        {
                            bool destinationBelow =
                                destination.Top >=
                                source.Top;

                            from =
                                new PointF(
                                    source.Left +
                                    source.Width /
                                    2f,
                                    destinationBelow
                                        ? source.Bottom
                                        : source.Top);

                            to =
                                new PointF(
                                    destination.Left +
                                    destination.Width /
                                    2f,
                                    destinationBelow
                                        ? destination.Top
                                        : destination.Bottom);
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

                        if (layoutKind == "hierarchy")
                        {
                            PointF[] routed =
                                DrawSmartArtElbowConnector(
                                    g,
                                    connector,
                                    from,
                                    to,
                                    positions,
                                    node.Id,
                                    node.Children[i],
                                    destinationIsAssistant,
                                    routedElbows);

                            if (routed != null)
                            {
                                routedElbows.Add(
                                    routed);
                            }
                        }
                        else if (layoutKind == "verticalProcess" ||
                                 (layoutKind == "process" &&
                                  !processSameRow))
                        {
                            PointF[] routed =
                                DrawSmartArtElbowConnector(
                                    g,
                                    connector,
                                    from,
                                    to,
                                    positions,
                                    node.Id,
                                    node.Children[i],
                                    false,
                                    routedElbows);

                            if (routed != null)
                            {
                                routedElbows.Add(
                                    routed);
                            }
                        }
                        else
                        {
                            g.DrawLine(
                                connector,
                                from,
                                to);
                        }
                    }
                }
            }
        }

        private static void DrawSmartArtElbowConnector(
            Graphics g,
            Pen connector,
            PointF from,
            PointF to,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId,
            bool preferVerticalChannel)
        {
            DrawSmartArtElbowConnector(
                g,
                connector,
                from,
                to,
                positions,
                sourceId,
                destinationId,
                preferVerticalChannel,
                null);
        }

        private static PointF[] DrawSmartArtElbowConnector(
            Graphics g,
            Pen connector,
            PointF from,
            PointF to,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId,
            bool preferVerticalChannel,
            List<PointF[]> existingRoutes)
        {
            PointF[] route =
                BuildSmartArtElbowRoute(
                    from,
                    to,
                    positions,
                    sourceId,
                    destinationId,
                    preferVerticalChannel,
                    existingRoutes);

            if (route != null &&
                route.Length >= 2)
            {
                g.DrawLines(
                    connector,
                    route);
            }

            return route;
        }

        private static PointF[] BuildSmartArtElbowRoute(
            PointF from,
            PointF to,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId,
            bool preferVerticalChannel)
        {
            return BuildSmartArtElbowRoute(
                from,
                to,
                positions,
                sourceId,
                destinationId,
                preferVerticalChannel,
                null);
        }

        private static PointF[] BuildSmartArtElbowRoute(
            PointF from,
            PointF to,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId,
            bool preferVerticalChannel,
            List<PointF[]> existingRoutes)
        {
            PointF[] route;

            if (preferVerticalChannel &&
                TryBuildSmartArtVerticalChannelRoute(
                    from,
                    to,
                    positions,
                    sourceId,
                    destinationId,
                    existingRoutes,
                    out route))
            {
                return route;
            }

            if (TryBuildSmartArtHorizontalChannelRoute(
                    from,
                    to,
                    positions,
                    sourceId,
                    destinationId,
                    existingRoutes,
                    out route))
            {
                return route;
            }

            if (!preferVerticalChannel &&
                TryBuildSmartArtVerticalChannelRoute(
                    from,
                    to,
                    positions,
                    sourceId,
                    destinationId,
                    existingRoutes,
                    out route))
            {
                return route;
            }

            float middleY =
                from.Y +
                (to.Y -
                 from.Y) /
                2f;

            return new PointF[]
            {
                from,
                new PointF(
                    from.X,
                    middleY),
                new PointF(
                    to.X,
                    middleY),
                to
            };
        }

        private static bool TryBuildSmartArtHorizontalChannelRoute(
            PointF from,
            PointF to,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId,
            List<PointF[]> existingRoutes,
            out PointF[] route)
        {
            float low =
                Math.Min(
                    from.Y,
                    to.Y);

            float high =
                Math.Max(
                    from.Y,
                    to.Y);

            float span =
                Math.Max(
                    0f,
                    high - low);

            float preferred =
                low +
                span /
                2f;

            List<float> candidates =
                new List<float>();

            candidates.Add(
                preferred);
            candidates.Add(
                low +
                span *
                0.25f);
            candidates.Add(
                low +
                span *
                0.75f);
            candidates.Add(
                low +
                Math.Min(
                    4f,
                    span /
                    2f));
            candidates.Add(
                high -
                Math.Min(
                    4f,
                    span /
                    2f));

            const float clearance =
                5f;

            if (positions != null)
            {
                foreach (KeyValuePair<string, RectangleF> pair
                    in positions)
                {
                    if (pair.Key == sourceId ||
                        pair.Key == destinationId)
                    {
                        continue;
                    }

                    candidates.Add(
                        pair.Value.Top -
                        clearance);
                    candidates.Add(
                        pair.Value.Bottom +
                        clearance);
                }
            }

            PointF[] bestRoute =
                null;
            double bestScore =
                double.MaxValue;
            float bestChannel =
                float.MaxValue;

            for (int i = 0;
                 i < candidates.Count;
                 i++)
            {
                float channel =
                    candidates[i];

                PointF[] candidate =
                    new PointF[]
                    {
                        from,
                        new PointF(
                            from.X,
                            channel),
                        new PointF(
                            to.X,
                            channel),
                        to
                    };

                if (!IsSmartArtRouteClear(
                        candidate,
                        positions,
                        sourceId,
                        destinationId))
                {
                    continue;
                }

                double score =
                    CalculateSmartArtRouteLength(
                        candidate) +
                    Math.Abs(
                        channel -
                        preferred) *
                    0.05 +
                    CalculateSmartArtRouteCrossingPenalty(
                        candidate,
                        existingRoutes);

                if (score <
                        bestScore -
                        0.001 ||
                    (Math.Abs(
                         score -
                         bestScore) <=
                     0.001 &&
                     channel <
                     bestChannel))
                {
                    bestRoute =
                        candidate;
                    bestScore =
                        score;
                    bestChannel =
                        channel;
                }
            }

            route =
                bestRoute;
            return route != null;
        }

        private static bool TryBuildSmartArtVerticalChannelRoute(
            PointF from,
            PointF to,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId,
            List<PointF[]> existingRoutes,
            out PointF[] route)
        {
            float low =
                Math.Min(
                    from.X,
                    to.X);

            float high =
                Math.Max(
                    from.X,
                    to.X);

            float span =
                Math.Max(
                    0f,
                    high - low);

            float preferred =
                low +
                span /
                2f;

            List<float> candidates =
                new List<float>();

            candidates.Add(
                preferred);
            candidates.Add(
                low +
                span *
                0.25f);
            candidates.Add(
                low +
                span *
                0.75f);

            const float clearance =
                5f;

            if (positions != null)
            {
                foreach (KeyValuePair<string, RectangleF> pair
                    in positions)
                {
                    if (pair.Key == sourceId ||
                        pair.Key == destinationId)
                    {
                        continue;
                    }

                    RectangleF obstacle =
                        pair.Value;

                    candidates.Add(
                        obstacle.Left -
                        clearance);
                    candidates.Add(
                        obstacle.Right +
                        clearance);
                }
            }

            PointF[] bestRoute =
                null;
            double bestScore =
                double.MaxValue;
            float bestChannel =
                float.MaxValue;

            for (int i = 0;
                 i < candidates.Count;
                 i++)
            {
                float channel =
                    candidates[i];

                PointF[] candidate =
                    new PointF[]
                    {
                        from,
                        new PointF(
                            channel,
                            from.Y),
                        new PointF(
                            channel,
                            to.Y),
                        to
                    };

                if (!IsSmartArtRouteClear(
                        candidate,
                        positions,
                        sourceId,
                        destinationId))
                {
                    continue;
                }

                double score =
                    CalculateSmartArtRouteLength(
                        candidate) +
                    Math.Abs(
                        channel -
                        preferred) *
                    0.05 +
                    CalculateSmartArtRouteCrossingPenalty(
                        candidate,
                        existingRoutes);

                if (score <
                        bestScore -
                        0.001 ||
                    (Math.Abs(
                         score -
                         bestScore) <=
                     0.001 &&
                     channel <
                     bestChannel))
                {
                    bestRoute =
                        candidate;
                    bestScore =
                        score;
                    bestChannel =
                        channel;
                }
            }

            route =
                bestRoute;
            return route != null;
        }

        private static double CalculateSmartArtRouteCrossingPenalty(
            PointF[] route,
            List<PointF[]> existingRoutes)
        {
            if (route == null ||
                existingRoutes == null ||
                existingRoutes.Count == 0)
            {
                return 0.0;
            }

            int crossings = 0;

            for (int r = 0;
                 r < existingRoutes.Count;
                 r++)
            {
                PointF[] existing =
                    existingRoutes[r];

                if (existing == null ||
                    existing.Length < 2)
                {
                    continue;
                }

                for (int i = 1;
                     i < route.Length;
                     i++)
                {
                    for (int j = 1;
                         j < existing.Length;
                         j++)
                    {
                        if (SmartArtSegmentsCrossInternally(
                                route[i - 1],
                                route[i],
                                existing[j - 1],
                                existing[j]))
                        {
                            crossings++;
                        }
                    }
                }
            }

            return crossings *
                10000.0;
        }

        private static bool SmartArtSegmentsCrossInternally(
            PointF a1,
            PointF a2,
            PointF b1,
            PointF b2)
        {
            const float epsilon =
                0.5f;

            bool aVertical =
                Math.Abs(
                    a1.X -
                    a2.X) <
                epsilon;
            bool bVertical =
                Math.Abs(
                    b1.X -
                    b2.X) <
                epsilon;

            if (aVertical ==
                bVertical)
            {
                return false;
            }

            PointF verticalStart =
                aVertical
                    ? a1
                    : b1;
            PointF verticalEnd =
                aVertical
                    ? a2
                    : b2;
            PointF horizontalStart =
                aVertical
                    ? b1
                    : a1;
            PointF horizontalEnd =
                aVertical
                    ? b2
                    : a2;

            float x =
                verticalStart.X;
            float y =
                horizontalStart.Y;

            float verticalMinY =
                Math.Min(
                    verticalStart.Y,
                    verticalEnd.Y);
            float verticalMaxY =
                Math.Max(
                    verticalStart.Y,
                    verticalEnd.Y);
            float horizontalMinX =
                Math.Min(
                    horizontalStart.X,
                    horizontalEnd.X);
            float horizontalMaxX =
                Math.Max(
                    horizontalStart.X,
                    horizontalEnd.X);

            return x >
                    horizontalMinX +
                    epsilon &&
                x <
                    horizontalMaxX -
                    epsilon &&
                y >
                    verticalMinY +
                    epsilon &&
                y <
                    verticalMaxY -
                    epsilon;
        }

        private static double CalculateSmartArtRouteLength(
            PointF[] route)
        {
            if (route == null ||
                route.Length < 2)
            {
                return double.MaxValue;
            }

            double length = 0.0;

            for (int i = 1;
                 i < route.Length;
                 i++)
            {
                length +=
                    Math.Abs(
                        route[i].X -
                        route[i - 1].X) +
                    Math.Abs(
                        route[i].Y -
                        route[i - 1].Y);
            }

            return length;
        }

        private static bool IsSmartArtRouteClear(
            PointF[] route,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId)
        {
            if (route == null ||
                route.Length < 2)
            {
                return false;
            }

            for (int i = 1;
                 i < route.Length;
                 i++)
            {
                if (!IsSmartArtSegmentClear(
                        route[i - 1],
                        route[i],
                        positions,
                        sourceId,
                        destinationId))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsSmartArtSegmentClear(
            PointF from,
            PointF to,
            Dictionary<string, RectangleF> positions,
            string sourceId,
            string destinationId)
        {
            if (positions == null)
                return true;

            const float clearance =
                3f;

            bool vertical =
                Math.Abs(
                    from.X -
                    to.X) < 0.5f;

            float minX =
                Math.Min(
                    from.X,
                    to.X);

            float maxX =
                Math.Max(
                    from.X,
                    to.X);

            float minY =
                Math.Min(
                    from.Y,
                    to.Y);

            float maxY =
                Math.Max(
                    from.Y,
                    to.Y);

            foreach (KeyValuePair<string, RectangleF> pair
                in positions)
            {
                if (pair.Key == sourceId ||
                    pair.Key == destinationId)
                {
                    continue;
                }

                RectangleF obstacle =
                    pair.Value;

                if (vertical)
                {
                    if (from.X <=
                            obstacle.Left -
                            clearance ||
                        from.X >=
                            obstacle.Right +
                            clearance)
                    {
                        continue;
                    }

                    if (maxY <=
                            obstacle.Top -
                            clearance ||
                        minY >=
                            obstacle.Bottom +
                            clearance)
                    {
                        continue;
                    }

                    return false;
                }

                if (from.Y <=
                        obstacle.Top -
                        clearance ||
                    from.Y >=
                        obstacle.Bottom +
                        clearance)
                {
                    continue;
                }

                if (maxX <=
                        obstacle.Left -
                        clearance ||
                    minX >=
                        obstacle.Right +
                        clearance)
                {
                    continue;
                }

                return false;
            }

            return true;
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
