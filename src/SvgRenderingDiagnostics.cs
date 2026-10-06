using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Xml;

namespace PptxViewer
{
    internal static partial class InternalPptxRenderer
    {
        public static void ValidateSvgRendering(
            string outputPath)
        {
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException(
                    "SVG diagnostic output path is required.",
                    "outputPath");

            string directory =
                Path.GetDirectoryName(
                    Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string svgText =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" viewBox=\"0 0 200 120\">" +
                "<defs>" +
                "<path id=\"reuseTriangle\" d=\"M 0 0 L 8 0 L 4 8 Z\"/>" +
                "<linearGradient id=\"g\" x1=\"0%\" y1=\"0%\" x2=\"100%\" y2=\"0%\" spreadMethod=\"reflect\" gradientTransform=\"rotate(22 .5 .5)\">" +
                "<stop offset=\"0%\" stop-color=\"#e84b4b\"/>" +
                "<stop offset=\"50%\" stop-color=\"#f2c94c\"/>" +
                "<stop offset=\"100%\" stop-color=\"#4b72e8\"/>" +
                "</linearGradient>" +
                "<clipPath id=\"clip\"><circle cx=\"100\" cy=\"78\" r=\"40\"/></clipPath>" +
                "<clipPath id=\"nestedClip\"><g transform=\"translate(22 0)\"><rect x=\"10\" y=\"8\" width=\"18\" height=\"16\"/></g></clipPath>" +
                "<clipPath id=\"evenoddClip\"><path d=\"M 176 2 H 196 V 22 H 176 Z M 181 7 H 191 V 17 H 181 Z\" clip-rule=\"evenodd\"/></clipPath>" +
                "<pattern id=\"pat\" patternUnits=\"userSpaceOnUse\" width=\"12\" height=\"12\">" +
                "<rect x=\"0\" y=\"0\" width=\"6\" height=\"12\" fill=\"#e85d75\"/>" +
                "<rect x=\"6\" y=\"0\" width=\"6\" height=\"12\" fill=\"#4c78d6\"/>" +
                "</pattern>" +
                "<mask id=\"mask\"><circle cx=\"158\" cy=\"88\" r=\"24\" fill=\"white\"/></mask>" +
                "<filter id=\"softBlur\"><feGaussianBlur in=\"SourceGraphic\" stdDeviation=\"3\"/></filter>" +
                "<filter id=\"dropShadow\"><feDropShadow in=\"SourceGraphic\" dx=\"2\" dy=\"1.5\" stdDeviation=\"1\" flood-color=\"#2244aa\" flood-opacity=\"0.7\"/></filter>" +
                "<filter id=\"offsetOnly\"><feOffset in=\"SourceGraphic\" dx=\"3\" dy=\"3\"/></filter>" +
                "</defs>" +
                "<g transform=\"matrix(1 0.10 -0.08 1 3 1)\"><rect x=\"16\" y=\"12\" width=\"58\" height=\"28\" rx=\"6\" fill=\"#20a77a\"/></g>" +
                "<use id=\"useTriangle\" xlink:href=\"#reuseTriangle\" x=\"134\" y=\"2\" color=\"hsl(326deg 53% 50% / 100%)\" fill=\"currentColor\"/>" +
                "<circle cx=\"154\" cy=\"28\" r=\"18\" fill=\"url(#g)\" transform=\"skewX(8)\"/>" +
                "<rect id=\"clipRuleTarget\" x=\"176\" y=\"2\" width=\"20\" height=\"20\" fill=\"#7b61ff\" clip-path=\"url(#evenoddClip)\"/>" +
                "<rect x=\"8\" y=\"6\" width=\"48\" height=\"22\" fill=\"#1677d2\" clip-path=\"url(#nestedClip)\"/>" +
                "<rect id=\"blurRect\" x=\"88\" y=\"14\" width=\"24\" height=\"18\" fill=\"#d62728\" filter=\"url(#softBlur)\"/>" +
                "<rect id=\"shadowRect\" x=\"78\" y=\"8\" width=\"6\" height=\"8\" fill=\"#f5c842\" filter=\"url(#dropShadow)\"/>" +
                "<rect id=\"offsetRect\" x=\"60\" y=\"2\" width=\"6\" height=\"6\" fill=\"#20b9c7\" filter=\"url(#offsetOnly)\"/>" +
                "<line id=\"dashLine\" x1=\"100\" y1=\"4\" x2=\"132\" y2=\"4\" stroke=\"#111\" stroke-width=\"2\" stroke-linejoin=\"miter\" stroke-miterlimit=\"2.5\" stroke-dasharray=\"4 3\" stroke-dashoffset=\"1\"/>" +
                "<path id=\"nonzeroPath\" d=\"M 2 30 H 14 V 38 H 2 Z M 5 32 H 11 V 36 H 5 Z\" fill=\"#f28c28\" fill-rule=\"nonzero\"/>" +
                "<path id=\"evenoddPath\" d=\"M 2 40 H 14 V 48 H 2 Z M 5 42 H 11 V 46 H 5 Z\" fill=\"#159a8c\" fill-rule=\"evenodd\"/>" +
                "<path d=\"M 12 52 C 30 38 42 68 60 52 S 90 38 108 52 Q 126 70 142 52 T 184 52\" fill=\"none\" stroke=\"#6f42a8\" stroke-width=\"2\"/>" +
                "<rect x=\"18\" y=\"66\" width=\"60\" height=\"42\" fill=\"url(#pat)\"/>" +
                "<rect x=\"124\" y=\"64\" width=\"68\" height=\"48\" fill=\"#29b36b\" mask=\"url(#mask)\"/>" +
                "<path d=\"M 18 88 A 82 48 0 0 1 182 88 L 182 116 L 18 116 Z\" " +
                "fill=\"url(#g)\" clip-path=\"url(#clip)\" stroke=\"#243447\" stroke-width=\"1.5\"/>" +
                "<text x=\"100\" y=\"63\" font-size=\"14\" fill=\"#20252b\">SVG</text>" +
                "</svg>";

            XmlDocument document =
                new XmlDocument();
            document.LoadXml(svgText);

            XmlNode reflectedGradient =
                FindSvgNodeById(
                    document,
                    "g");

            using (Brush spreadBrush =
                CreateSvgGradientBrush(
                    reflectedGradient,
                    document,
                    new RectangleF(
                        0f,
                        0f,
                        200f,
                        120f),
                    new RectangleF(
                        0f,
                        0f,
                        800f,
                        480f),
                    0f,
                    0f,
                    4f,
                    4f,
                    1f))
            {
                LinearGradientBrush reflected =
                    spreadBrush as LinearGradientBrush;

                if (reflected == null ||
                    reflected.WrapMode !=
                        WrapMode.TileFlipXY)
                {
                    throw new InvalidOperationException(
                        "SVG reflect spreadMethod was not applied.");
                }
            }

            XmlNode blurRect =
                FindSvgNodeById(
                    document,
                    "blurRect");

            float parsedBlurX;
            float parsedBlurY;

            if (!TryReadSvgGaussianBlur(
                    blurRect,
                    document,
                    4f,
                    4f,
                    out parsedBlurX,
                    out parsedBlurY) ||
                Math.Abs(
                    parsedBlurX -
                    12f) > 0.1f ||
                Math.Abs(
                    parsedBlurY -
                    12f) > 0.1f)
            {
                throw new InvalidOperationException(
                    "SVG feGaussianBlur stdDeviation was not parsed correctly.");
            }

            XmlNode shadowRect =
                FindSvgNodeById(
                    document,
                    "shadowRect");

            float shadowOffsetX;
            float shadowOffsetY;
            float shadowBlurX;
            float shadowBlurY;
            Color parsedShadowColor;

            if (!TryReadSvgDropShadow(
                    shadowRect,
                    document,
                    4f,
                    4f,
                    out shadowOffsetX,
                    out shadowOffsetY,
                    out shadowBlurX,
                    out shadowBlurY,
                    out parsedShadowColor) ||
                Math.Abs(
                    shadowOffsetX -
                    8f) > 0.1f ||
                Math.Abs(
                    shadowOffsetY -
                    6f) > 0.1f ||
                Math.Abs(
                    shadowBlurX -
                    4f) > 0.1f ||
                Math.Abs(
                    shadowBlurY -
                    4f) > 0.1f ||
                parsedShadowColor.B <
                    130 ||
                parsedShadowColor.A <
                    150)
            {
                throw new InvalidOperationException(
                    "SVG feDropShadow parameters were not parsed correctly.");
            }

            XmlNode offsetRect =
                FindSvgNodeById(
                    document,
                    "offsetRect");

            float parsedOffsetX;
            float parsedOffsetY;

            if (!TryReadStandaloneSvgOffset(
                    offsetRect,
                    document,
                    4f,
                    4f,
                    out parsedOffsetX,
                    out parsedOffsetY) ||
                Math.Abs(
                    parsedOffsetX -
                    12f) > 0.1f ||
                Math.Abs(
                    parsedOffsetY -
                    12f) > 0.1f)
            {
                throw new InvalidOperationException(
                    "SVG feOffset parameters were not parsed correctly.");
            }

            XmlNode dashLine =
                FindSvgNodeById(
                    document,
                    "dashLine");

            using (Pen dashPen =
                new Pen(
                    Color.Black,
                    8f))
            {
                ApplySvgStrokeMiterLimit(
                    dashPen,
                    dashLine);

                if (Math.Abs(
                        dashPen.MiterLimit -
                        2.5f) > 0.05f ||
                    !ApplySvgStrokeDashPattern(
                        dashPen,
                        dashLine,
                        new RectangleF(
                            0f,
                            0f,
                            800f,
                            480f),
                        4f,
                        4f,
                        8f) ||
                    dashPen.DashStyle !=
                        DashStyle.Custom ||
                    dashPen.DashPattern == null ||
                    dashPen.DashPattern.Length !=
                        2 ||
                    Math.Abs(
                        dashPen.DashPattern[0] -
                        2f) > 0.05f ||
                    Math.Abs(
                        dashPen.DashPattern[1] -
                        1.5f) > 0.05f ||
                    Math.Abs(
                        dashPen.DashOffset -
                        0.5f) > 0.05f)
                {
                    throw new InvalidOperationException(
                        "SVG stroke dash pattern was not configured correctly.");
                }
            }

            XmlNode nonzeroNode =
                FindSvgNodeById(
                    document,
                    "nonzeroPath");
            XmlNode evenoddNode =
                FindSvgNodeById(
                    document,
                    "evenoddPath");

            using (GraphicsPath nonzeroGeometry =
                BuildEnhancedSvgElementPath(
                    nonzeroNode,
                    new RectangleF(
                        0f,
                        0f,
                        800f,
                        480f),
                    0f,
                    0f,
                    4f,
                    4f))
            using (GraphicsPath evenoddGeometry =
                BuildEnhancedSvgElementPath(
                    evenoddNode,
                    new RectangleF(
                        0f,
                        0f,
                        800f,
                        480f),
                    0f,
                    0f,
                    4f,
                    4f))
            {
                ApplySvgFillRule(
                    nonzeroGeometry,
                    nonzeroNode);
                ApplySvgFillRule(
                    evenoddGeometry,
                    evenoddNode);

                if (nonzeroGeometry == null ||
                    evenoddGeometry == null ||
                    nonzeroGeometry.FillMode !=
                        FillMode.Winding ||
                    evenoddGeometry.FillMode !=
                        FillMode.Alternate)
                {
                    throw new InvalidOperationException(
                        "SVG fill-rule was not mapped to the expected GDI+ FillMode.");
                }
            }

            XmlNode clipRuleTarget =
                FindSvgNodeById(
                    document,
                    "clipRuleTarget");

            using (GraphicsPath evenoddClip =
                BuildSvgClipPath(
                    clipRuleTarget,
                    document,
                    new RectangleF(
                        704f,
                        8f,
                        80f,
                        80f),
                    new RectangleF(
                        0f,
                        0f,
                        800f,
                        480f),
                    0f,
                    0f,
                    4f,
                    4f))
            {
                if (evenoddClip == null ||
                    evenoddClip.FillMode !=
                        FillMode.Alternate ||
                    !evenoddClip.IsVisible(
                        712f,
                        16f) ||
                    evenoddClip.IsVisible(
                        744f,
                        48f))
                {
                    throw new InvalidOperationException(
                        "SVG clip-rule evenodd was not preserved in the clip geometry.");
                }
            }

            using (Pen defaultMiterPen =
                new Pen(
                    Color.Black,
                    4f))
            {
                ApplySvgStrokeMiterLimit(
                    defaultMiterPen,
                    nonzeroNode);

                if (Math.Abs(
                        defaultMiterPen.MiterLimit -
                        4f) > 0.05f)
                {
                    throw new InvalidOperationException(
                        "SVG default stroke-miterlimit was not set to 4.");
                }
            }

            XmlNode useTriangle =
                FindSvgNodeById(
                    document,
                    "useTriangle");

            Color currentColorFill =
                ReadSvgColorInherited(
                    useTriangle,
                    "fill",
                    Color.Transparent);

            Color shortHexAlpha =
                ParseSvgColorValue(
                    "#c388",
                    Color.Transparent);

            if (shortHexAlpha.R != 204 ||
                shortHexAlpha.G != 51 ||
                shortHexAlpha.B != 136 ||
                shortHexAlpha.A != 136)
            {
                throw new InvalidOperationException(
                    "SVG #RGBA shorthand color was not parsed correctly.");
            }

            Color legacyRgba =
                ParseSvgColorValue(
                    "rgba(20%, 40%, 80%, 50%)",
                    Color.Transparent);

            Color hslAlpha =
                ParseSvgColorValue(
                    "hsla(210, 60%, 40%, 0.5)",
                    Color.Transparent);

            if (Math.Abs(
                    hslAlpha.R -
                    41) > 2 ||
                Math.Abs(
                    hslAlpha.G -
                    102) > 2 ||
                Math.Abs(
                    hslAlpha.B -
                    163) > 2 ||
                Math.Abs(
                    hslAlpha.A -
                    128) > 1)
            {
                throw new InvalidOperationException(
                    "SVG hsla color syntax was not parsed correctly.");
            }

            if (Math.Abs(
                    legacyRgba.R -
                    51) > 1 ||
                Math.Abs(
                    legacyRgba.G -
                    102) > 1 ||
                Math.Abs(
                    legacyRgba.B -
                    204) > 1 ||
                Math.Abs(
                    legacyRgba.A -
                    128) > 1)
            {
                throw new InvalidOperationException(
                    "SVG rgba percentage color syntax was not parsed correctly.");
            }

            if (currentColorFill.R < 180 ||
                currentColorFill.B < 120 ||
                currentColorFill.G > 100)
            {
                throw new InvalidOperationException(
                    "SVG currentColor did not resolve from the inherited color property.");
            }

            using (GraphicsPath reusedGeometry =
                BuildEnhancedSvgElementPath(
                    useTriangle,
                    new RectangleF(
                        0f,
                        0f,
                        800f,
                        480f),
                    0f,
                    0f,
                    4f,
                    4f))
            {
                if (reusedGeometry == null ||
                    reusedGeometry.PointCount <
                        3)
                {
                    throw new InvalidOperationException(
                        "SVG use geometry reference did not resolve.");
                }

                RectangleF reusedBounds =
                    reusedGeometry.GetBounds();

                if (Math.Abs(
                        reusedBounds.X -
                        536f) > 0.5f ||
                    Math.Abs(
                        reusedBounds.Y -
                        8f) > 0.5f ||
                    Math.Abs(
                        reusedBounds.Width -
                        32f) > 0.5f ||
                    Math.Abs(
                        reusedBounds.Height -
                        32f) > 0.5f)
                {
                    throw new InvalidOperationException(
                        "SVG use x/y translation was not applied to referenced geometry.");
                }
            }

            using (Bitmap bitmap =
                new Bitmap(
                    800,
                    480,
                    PixelFormat.Format32bppArgb))
            using (Graphics graphics =
                Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode =
                    System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                bool drew =
                    DrawEnhancedSvgChildren(
                        graphics,
                        document.DocumentElement,
                        new RectangleF(
                            0f,
                            0f,
                            800f,
                            480f),
                        0f,
                        0f,
                        4f,
                        4f);

                if (!drew)
                {
                    throw new InvalidOperationException(
                        "Enhanced SVG renderer did not draw the synthetic document.");
                }

                int colored = 0;
                HashSet<int> colors =
                    new HashSet<int>();

                for (int y = 0;
                     y < bitmap.Height;
                     y += 8)
                {
                    for (int x = 0;
                         x < bitmap.Width;
                         x += 8)
                    {
                        Color color =
                            bitmap.GetPixel(
                                x,
                                y);

                        if (color.R > 248 &&
                            color.G > 248 &&
                            color.B > 248)
                        {
                            continue;
                        }

                        colored++;

                        int bucket =
                            ((color.R / 32) << 10) |
                            ((color.G / 32) << 5) |
                            (color.B / 32);

                        colors.Add(bucket);
                    }
                }

                if (colored < 120)
                {
                    throw new InvalidOperationException(
                        "Synthetic SVG output contains too little rendered content.");
                }

                if (colors.Count < 6)
                {
                    throw new InvalidOperationException(
                        "Synthetic SVG gradient/mixed-color rendering was not preserved.");
                }

                Color nestedClipOutside =
                    bitmap.GetPixel(
                        60,
                        56);
                Color nestedClipInside =
                    bitmap.GetPixel(
                        144,
                        56);

                if (nestedClipOutside.R < 245 ||
                    nestedClipOutside.G < 245 ||
                    nestedClipOutside.B < 245 ||
                    nestedClipInside.B < 140 ||
                    nestedClipInside.R > 100)
                {
                    throw new InvalidOperationException(
                        "Nested transformed SVG clipPath geometry was not positioned correctly.");
                }

                Color useTrianglePixel =
                    bitmap.GetPixel(
                        552,
                        20);

                if (useTrianglePixel.R <
                        130 ||
                    useTrianglePixel.B <
                        90 ||
                    useTrianglePixel.G >
                        120)
                {
                    throw new InvalidOperationException(
                        "SVG use referenced geometry did not render with the use-side fill.");
                }

                Color clipRuleOuter =
                    bitmap.GetPixel(
                        712,
                        16);
                Color clipRuleHole =
                    bitmap.GetPixel(
                        744,
                        48);

                bool clipOuterVisible =
                    clipRuleOuter.B >
                        150 &&
                    clipRuleOuter.R >
                        70 &&
                    clipRuleOuter.G <
                        150;

                bool clipHoleEmpty =
                    clipRuleHole.R >
                        245 &&
                    clipRuleHole.G >
                        245 &&
                    clipRuleHole.B >
                        245;

                if (!clipOuterVisible ||
                    !clipHoleEmpty)
                {
                    throw new InvalidOperationException(
                        "SVG clip-rule evenodd did not preserve the expected clipped hole.");
                }

                Color nonzeroCenter =
                    bitmap.GetPixel(
                        32,
                        136);
                Color evenoddCenter =
                    bitmap.GetPixel(
                        32,
                        176);

                bool nonzeroFilled =
                    nonzeroCenter.R >
                        180 &&
                    nonzeroCenter.G >
                        70 &&
                    nonzeroCenter.B <
                        100;

                bool evenoddHole =
                    evenoddCenter.R >
                        245 &&
                    evenoddCenter.G >
                        245 &&
                    evenoddCenter.B >
                        245;

                if (!nonzeroFilled ||
                    !evenoddHole)
                {
                    throw new InvalidOperationException(
                        "SVG fill-rule nonzero/evenodd rendering did not preserve the expected inner region.");
                }

                int dashDarkSamples =
                    0;
                int dashLightSamples =
                    0;

                for (int x = 400;
                     x <= 528;
                     x += 2)
                {
                    Color dashPixel =
                        bitmap.GetPixel(
                            x,
                            16);

                    if (dashPixel.R < 100 &&
                        dashPixel.G < 100 &&
                        dashPixel.B < 100)
                    {
                        dashDarkSamples++;
                    }
                    else if (dashPixel.R > 240 &&
                        dashPixel.G > 240 &&
                        dashPixel.B > 240)
                    {
                        dashLightSamples++;
                    }
                }

                if (dashDarkSamples < 8 ||
                    dashLightSamples < 6)
                {
                    throw new InvalidOperationException(
                        "SVG stroke-dasharray did not render visible dash and gap segments.");
                }

                Color originalOffsetPosition =
                    bitmap.GetPixel(
                        244,
                        12);
                Color shiftedOffsetPosition =
                    bitmap.GetPixel(
                        272,
                        40);

                bool originalLooksEmpty =
                    originalOffsetPosition.R >
                        245 &&
                    originalOffsetPosition.G >
                        245 &&
                    originalOffsetPosition.B >
                        245;

                bool shiftedLooksCyan =
                    shiftedOffsetPosition.G >
                        125 &&
                    shiftedOffsetPosition.B >
                        130 &&
                    shiftedOffsetPosition.R <
                        120;

                if (!originalLooksEmpty ||
                    !shiftedLooksCyan)
                {
                    throw new InvalidOperationException(
                        "SVG feOffset approximation did not move the source geometry.");
                }

                bool foundDropShadow =
                    false;

                for (int y = 42;
                     y <= 72 &&
                     !foundDropShadow;
                     y += 2)
                {
                    for (int x = 336;
                         x <= 352;
                         x += 2)
                    {
                        Color shadowPixel =
                            bitmap.GetPixel(
                                x,
                                y);

                        if (shadowPixel.B -
                                shadowPixel.R >= 8 &&
                            shadowPixel.B -
                                shadowPixel.G >= 3 &&
                            shadowPixel.B < 252)
                        {
                            foundDropShadow =
                                true;
                            break;
                        }
                    }
                }

                if (!foundDropShadow)
                {
                    throw new InvalidOperationException(
                        "SVG feDropShadow approximation did not render the configured shadow.");
                }

                bool foundBlurHalo =
                    false;

                for (int y = 76;
                     y <= 112 &&
                     !foundBlurHalo;
                     y += 4)
                {
                    for (int x = 336;
                         x < 352;
                         x += 2)
                    {
                        Color blurHalo =
                            bitmap.GetPixel(
                                x,
                                y);

                        if (blurHalo.R -
                                blurHalo.G >= 3 &&
                            blurHalo.R -
                                blurHalo.B >= 3 &&
                            blurHalo.G < 253)
                        {
                            foundBlurHalo =
                                true;
                            break;
                        }
                    }
                }

                if (!foundBlurHalo)
                {
                    throw new InvalidOperationException(
                        "SVG feGaussianBlur approximation did not render outside the source geometry.");
                }

                Color patternA =
                    bitmap.GetPixel(
                        90,
                        300);
                Color patternB =
                    bitmap.GetPixel(
                        114,
                        300);

                if (Math.Abs(
                        patternA.R -
                        patternB.R) < 20 &&
                    Math.Abs(
                        patternA.B -
                        patternB.B) < 20)
                {
                    throw new InvalidOperationException(
                        "SVG pattern fill did not repeat distinct tile colors.");
                }

                Color maskCenter =
                    bitmap.GetPixel(
                        632,
                        352);
                // Pick a point inside the masked source rectangle but
                // outside both the mask circle and the later clipped gradient path.
                Color maskCorner =
                    bitmap.GetPixel(
                        752,
                        272);

                if ((maskCenter.R > 245 &&
                     maskCenter.G > 245 &&
                     maskCenter.B > 245) ||
                    maskCorner.R < 245 ||
                    maskCorner.G < 245 ||
                    maskCorner.B < 245)
                {
                    throw new InvalidOperationException(
                        "SVG mask clipping did not preserve the expected center/outside pixels.");
                }

                bitmap.Save(
                    outputPath,
                    ImageFormat.Png);
            }

            FileInfo info =
                new FileInfo(outputPath);

            if (!info.Exists ||
                info.Length <= 0)
            {
                throw new InvalidOperationException(
                    "SVG diagnostic PNG was not created.");
            }
        }
    }
}
