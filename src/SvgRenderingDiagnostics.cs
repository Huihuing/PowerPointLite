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
                "<path id=\"styledReference\" d=\"M 0 0 H 7 V 7 H 0 Z\" fill=\"#2f8f5b\" stroke=\"#163a27\" stroke-width=\"1.5\"/>" +
                "<symbol id=\"reuseSymbol\" viewBox=\"0 0 10 10\"><rect x=\"0\" y=\"0\" width=\"6\" height=\"6\"/><circle cx=\"8\" cy=\"3\" r=\"2\"/></symbol>" +
                "<symbol id=\"styledSymbol\" fill=\"#6b4fd3\" stroke=\"#2b1c68\"><rect x=\"0\" y=\"0\" width=\"5\" height=\"5\"/></symbol>" +
                "<g id=\"styledGroup\" fill=\"#cc6b32\" stroke-width=\"2.25\"><rect x=\"0\" y=\"0\" width=\"4\" height=\"4\"/></g>" +
                "<g id=\"multiStyleGroup\"><rect x=\"0\" y=\"0\" width=\"6\" height=\"8\" fill=\"#d43c32\"/><rect x=\"7\" y=\"0\" width=\"6\" height=\"8\" fill=\"#2864c8\"/></g>" +
                "<use id=\"nestedUseSource\" xlink:href=\"#reuseSymbol\" x=\"2\" y=\"1\" width=\"20\" height=\"10\" preserveAspectRatio=\"none\"/>" +
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
                "<filter id=\"offsetBlurChain\"><feOffset in=\"SourceGraphic\" dx=\"8\" dy=\"2\" result=\"shifted\"/><feGaussianBlur in=\"shifted\" stdDeviation=\"1\"/></filter>" +
                "<filter id=\"offsetBlurBlend\"><feOffset in=\"SourceGraphic\" dx=\"8\" dy=\"2\" result=\"shifted2\"/><feGaussianBlur in=\"shifted2\" stdDeviation=\"1\" result=\"blurred2\"/><feBlend in=\"blurred2\" in2=\"SourceGraphic\" mode=\"normal\"/></filter>" +
                "<filter id=\"swapRedBlue\"><feColorMatrix in=\"SourceGraphic\" type=\"matrix\" values=\"0 0 1 0 0  0 1 0 0 0  1 0 0 0 0  0 0 0 1 0\"/></filter>" +
                "<filter id=\"desaturate\"><feColorMatrix in=\"SourceGraphic\" type=\"saturate\" values=\"0\"/></filter>" +
                "<filter id=\"hueRotate\"><feColorMatrix in=\"SourceGraphic\" type=\"hueRotate\" values=\"240\"/></filter>" +
                "<filter id=\"lumaAlpha\"><feColorMatrix in=\"SourceGraphic\" type=\"luminanceToAlpha\"/></filter>" +
                "<rect id=\"offsetBlurBlendRect\" x=\"2\" y=\"12\" width=\"8\" height=\"6\" fill=\"#d64545\" filter=\"url(#offsetBlurBlend)\"/>" +
                "</defs>" +
                "<g transform=\"matrix(1 0.10 -0.08 1 3 1)\"><rect x=\"16\" y=\"12\" width=\"58\" height=\"28\" rx=\"6\" fill=\"#20a77a\"/></g>" +
                "<g color=\"hsl(326deg 53% 50% / 100%)\"><use id=\"useTriangle\" xlink:href=\"#reuseTriangle\" x=\"134\" y=\"2\" color=\"inherit\" fill=\"currentColor\"/></g>" +
                "<use id=\"styledUse\" xlink:href=\"#styledReference\" x=\"118\" y=\"2\"/>" +
                "<use id=\"styledSymbolUse\" xlink:href=\"#styledSymbol\" x=\"106\" y=\"2\"/>" +
                "<use id=\"styledGroupUse\" xlink:href=\"#styledGroup\" x=\"98\" y=\"2\"/>" +
                "<use id=\"multiStyleUse\" xlink:href=\"#multiStyleGroup\" x=\"2\" y=\"2\"/>" +
                "<use id=\"useSymbol\" xlink:href=\"#nestedUseSource\" x=\"146\" y=\"4\" fill=\"#3a7bd5\"/>" +
                "<circle cx=\"154\" cy=\"28\" r=\"18\" fill=\"url(#g)\" transform=\"skewX(8)\"/>" +
                "<rect id=\"clipRuleTarget\" x=\"176\" y=\"2\" width=\"20\" height=\"20\" fill=\"#7b61ff\" clip-path=\"url(#evenoddClip)\"/>" +
                "<rect x=\"8\" y=\"6\" width=\"48\" height=\"22\" fill=\"#1677d2\" clip-path=\"url(#nestedClip)\"/>" +
                "<rect id=\"blurRect\" x=\"88\" y=\"14\" width=\"24\" height=\"18\" fill=\"#d62728\" filter=\"url(#softBlur)\"/>" +
                "<rect id=\"shadowRect\" x=\"78\" y=\"8\" width=\"6\" height=\"8\" fill=\"#f5c842\" filter=\"url(#dropShadow)\"/>" +
                "<rect id=\"offsetRect\" x=\"60\" y=\"2\" width=\"6\" height=\"6\" fill=\"#20b9c7\" filter=\"url(#offsetOnly)\"/>" +
                "<rect id=\"offsetBlurChainRect\" x=\"2\" y=\"2\" width=\"8\" height=\"6\" fill=\"#d64545\" filter=\"url(#offsetBlurChain)\"/>" +
                "<rect id=\"colorMatrixRect\" x=\"14\" y=\"2\" width=\"8\" height=\"6\" fill=\"#e04030\" filter=\"url(#swapRedBlue)\"/>" +
                "<rect id=\"saturateRect\" x=\"24\" y=\"2\" width=\"8\" height=\"6\" fill=\"#e04030\" filter=\"url(#desaturate)\"/>" +
                "<rect id=\"hueRotateRect\" x=\"34\" y=\"2\" width=\"8\" height=\"6\" fill=\"#ff0000\" filter=\"url(#hueRotate)\"/>" +
                "<rect id=\"lumaAlphaRect\" x=\"44\" y=\"2\" width=\"8\" height=\"6\" fill=\"#ffffff\" filter=\"url(#lumaAlpha)\"/>" +
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

            XmlNode styledUseNode =
                FindSvgNodeById(
                    document,
                    "styledUse");

            string referencedFill =
                GetSvgStyleInherited(
                    styledUseNode,
                    "fill");
            string referencedStroke =
                GetSvgStyleInherited(
                    styledUseNode,
                    "stroke");

            if (!string.Equals(
                    referencedFill,
                    "#2f8f5b",
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    referencedStroke,
                    "#163a27",
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "SVG direct use did not inherit explicit presentation styles from its referenced primitive.");
            }

            XmlNode styledSymbolUseNode =
                FindSvgNodeById(
                    document,
                    "styledSymbolUse");
            XmlNode styledGroupUseNode =
                FindSvgNodeById(
                    document,
                    "styledGroupUse");

            if (!string.Equals(
                    GetSvgStyleInherited(
                        styledSymbolUseNode,
                        "fill"),
                    "#6b4fd3",
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    GetSvgStyleInherited(
                        styledSymbolUseNode,
                        "stroke"),
                    "#2b1c68",
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    GetSvgStyleInherited(
                        styledGroupUseNode,
                        "fill"),
                    "#cc6b32",
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    GetSvgStyleInherited(
                        styledGroupUseNode,
                        "stroke-width"),
                    "2.25",
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "SVG use did not inherit container-level presentation styles from symbol or group references.");
            }

            XmlNode multiStyleUseNode =
                FindSvgNodeById(
                    document,
                    "multiStyleUse");

            using (Bitmap useStyleBitmap =
                new Bitmap(
                    80,
                    48,
                    PixelFormat.Format32bppArgb))
            using (Graphics useStyleGraphics =
                Graphics.FromImage(
                    useStyleBitmap))
            {
                useStyleGraphics.Clear(
                    Color.White);
                useStyleGraphics.SmoothingMode =
                    SmoothingMode.None;

                if (!TryDrawEnhancedSvgUseContainer(
                        useStyleGraphics,
                        multiStyleUseNode,
                        new RectangleF(
                            0f,
                            0f,
                            80f,
                            48f),
                        0f,
                        0f,
                        4f,
                        4f))
                {
                    throw new InvalidOperationException(
                        "SVG use container expansion did not render the referenced group.");
                }

                Color firstChild =
                    useStyleBitmap.GetPixel(
                        12,
                        16);
                Color secondChild =
                    useStyleBitmap.GetPixel(
                        40,
                        16);

                bool firstRed =
                    firstChild.R >
                        150 &&
                    firstChild.G <
                        110 &&
                    firstChild.B <
                        100;
                bool secondBlue =
                    secondChild.B >
                        130 &&
                    secondChild.R <
                        100;

                if (!firstRed ||
                    !secondBlue)
                {
                    throw new InvalidOperationException(
                        "SVG use container expansion did not preserve distinct child fill styles.");
                }
            }

            XmlNode nestedUseNode =
                FindSvgNodeById(
                    document,
                    "useSymbol");

            using (GraphicsPath nestedUsePath =
                BuildEnhancedSvgElementPath(
                    nestedUseNode,
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
                if (nestedUsePath == null ||
                    nestedUsePath.PointCount == 0)
                {
                    throw new InvalidOperationException(
                        "SVG symbol or nested use reference did not produce geometry.");
                }

                RectangleF nestedBounds =
                    nestedUsePath.GetBounds();

                if (nestedBounds.Width < 65f ||
                    nestedBounds.Height < 20f)
                {
                    throw new InvalidOperationException(
                        "SVG symbol viewBox was not scaled by use width and height.");
                }
            }

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

            XmlNode offsetBlurChainRect =
                FindSvgNodeById(
                    document,
                    "offsetBlurChainRect");

            float chainOffsetX;
            float chainOffsetY;
            float chainBlurX;
            float chainBlurY;
            bool chainBlendSource;

            if (!TryReadSvgOffsetGaussianChain(
                    offsetBlurChainRect,
                    document,
                    4f,
                    4f,
                    out chainOffsetX,
                    out chainOffsetY,
                    out chainBlurX,
                    out chainBlurY,
                    out chainBlendSource) ||
                Math.Abs(
                    chainOffsetX -
                    32f) > 0.1f ||
                Math.Abs(
                    chainOffsetY -
                    8f) > 0.1f ||
                Math.Abs(
                    chainBlurX -
                    4f) > 0.1f ||
                Math.Abs(
                    chainBlurY -
                    4f) > 0.1f ||
                chainBlendSource)
            {
                throw new InvalidOperationException(
                    "SVG feOffset/feGaussianBlur filter chain was not parsed correctly.");
            }

            using (GraphicsPath chainPath =
                BuildEnhancedSvgElementPath(
                    offsetBlurChainRect,
                    new RectangleF(
                        0f,
                        0f,
                        160f,
                        80f),
                    0f,
                    0f,
                    4f,
                    4f))
            using (Bitmap chainBitmap =
                new Bitmap(
                    160,
                    80,
                    PixelFormat.Format32bppArgb))
            using (Graphics chainGraphics =
                Graphics.FromImage(
                    chainBitmap))
            {
                chainGraphics.Clear(
                    Color.White);

                if (!DrawSvgOffsetGaussianChainApproximation(
                        chainGraphics,
                        offsetBlurChainRect,
                        document,
                        chainPath,
                        4f,
                        4f))
                {
                    throw new InvalidOperationException(
                        "SVG feOffset/feGaussianBlur filter chain was not rendered.");
                }

                Color originalCenter =
                    chainBitmap.GetPixel(
                        24,
                        20);
                Color shiftedCenter =
                    chainBitmap.GetPixel(
                        56,
                        28);

                bool originalMostlyEmpty =
                    originalCenter.R >
                        235 &&
                    originalCenter.G >
                        235 &&
                    originalCenter.B >
                        235;
                bool shiftedVisible =
                    shiftedCenter.R >
                        shiftedCenter.G +
                        25 &&
                    shiftedCenter.R >
                        shiftedCenter.B +
                        25 &&
                    shiftedCenter.R <
                        250;

                if (!originalMostlyEmpty ||
                    !shiftedVisible)
                {
                    throw new InvalidOperationException(
                        "SVG offset/blur chain did not produce only the shifted blurred result.");
                }
            }

            XmlNode offsetBlurBlendRect =
                FindSvgNodeById(
                    document,
                    "offsetBlurBlendRect");

            float blendOffsetX;
            float blendOffsetY;
            float blendBlurX;
            float blendBlurY;
            bool blendSourceGraphic;

            if (!TryReadSvgOffsetGaussianChain(
                    offsetBlurBlendRect,
                    document,
                    4f,
                    4f,
                    out blendOffsetX,
                    out blendOffsetY,
                    out blendBlurX,
                    out blendBlurY,
                    out blendSourceGraphic) ||
                !blendSourceGraphic)
            {
                throw new InvalidOperationException(
                    "SVG feOffset/feGaussianBlur/feBlend filter chain was not parsed correctly.");
            }

            using (GraphicsPath blendPath =
                BuildEnhancedSvgElementPath(
                    offsetBlurBlendRect,
                    new RectangleF(
                        0f,
                        0f,
                        160f,
                        100f),
                    0f,
                    0f,
                    4f,
                    4f))
            using (Bitmap blendBitmap =
                new Bitmap(
                    160,
                    100,
                    PixelFormat.Format32bppArgb))
            using (Graphics blendGraphics =
                Graphics.FromImage(
                    blendBitmap))
            {
                blendGraphics.Clear(
                    Color.White);

                if (!DrawSvgOffsetGaussianChainApproximation(
                        blendGraphics,
                        offsetBlurBlendRect,
                        document,
                        blendPath,
                        4f,
                        4f))
                {
                    throw new InvalidOperationException(
                        "SVG feBlend filter chain was not rendered.");
                }

                Color originalCenter =
                    blendBitmap.GetPixel(
                        24,
                        60);
                Color shiftedCenter =
                    blendBitmap.GetPixel(
                        56,
                        68);

                bool originalVisible =
                    originalCenter.R >
                        originalCenter.G +
                        25 &&
                    originalCenter.R >
                        originalCenter.B +
                        25 &&
                    originalCenter.R <
                        250;
                bool shiftedVisible =
                    shiftedCenter.R >
                        shiftedCenter.G +
                        25 &&
                    shiftedCenter.R >
                        shiftedCenter.B +
                        25 &&
                    shiftedCenter.R <
                        250;

                if (!originalVisible ||
                    !shiftedVisible)
                {
                    throw new InvalidOperationException(
                        "SVG normal feBlend chain did not preserve SourceGraphic with the filtered result.");
                }
            }

            XmlNode colorMatrixRect =
                FindSvgNodeById(
                    document,
                    "colorMatrixRect");
            XmlNode saturateRect =
                FindSvgNodeById(
                    document,
                    "saturateRect");

            float[] swapMatrix;
            float[] saturationMatrix;

            if (!TryReadSvgColorMatrix(
                    colorMatrixRect,
                    document,
                    out swapMatrix) ||
                !TryReadSvgColorMatrix(
                    saturateRect,
                    document,
                    out saturationMatrix))
            {
                throw new InvalidOperationException(
                    "SVG feColorMatrix matrix or saturate mode was not parsed.");
            }

            Color swapped =
                ApplySvgColorMatrix(
                    Color.FromArgb(
                        255,
                        224,
                        64,
                        48),
                    swapMatrix);
            Color desaturated =
                ApplySvgColorMatrix(
                    Color.FromArgb(
                        255,
                        224,
                        64,
                        48),
                    saturationMatrix);

            if (swapped.B < 210 ||
                swapped.R > 70 ||
                Math.Abs(
                    desaturated.R -
                    desaturated.G) > 2 ||
                Math.Abs(
                    desaturated.G -
                    desaturated.B) > 2)
            {
                throw new InvalidOperationException(
                    "SVG feColorMatrix did not transform colors as expected.");
            }

            XmlNode hueRotateRect =
                FindSvgNodeById(
                    document,
                    "hueRotateRect");
            XmlNode lumaAlphaRect =
                FindSvgNodeById(
                    document,
                    "lumaAlphaRect");

            float[] hueMatrix;
            float[] lumaMatrix;

            if (!TryReadSvgColorMatrix(
                    hueRotateRect,
                    document,
                    out hueMatrix) ||
                !TryReadSvgColorMatrix(
                    lumaAlphaRect,
                    document,
                    out lumaMatrix))
            {
                throw new InvalidOperationException(
                    "SVG feColorMatrix hueRotate or luminanceToAlpha mode was not parsed.");
            }

            Color hueRotated =
                ApplySvgColorMatrix(
                    Color.Red,
                    hueMatrix);
            Color luminanceAlpha =
                ApplySvgColorMatrix(
                    Color.White,
                    lumaMatrix);

            if (hueRotated.B < 220 ||
                hueRotated.R > 40 ||
                luminanceAlpha.A < 250 ||
                luminanceAlpha.R > 3 ||
                luminanceAlpha.G > 3 ||
                luminanceAlpha.B > 3)
            {
                throw new InvalidOperationException(
                    "SVG feColorMatrix hueRotate or luminanceToAlpha output was incorrect.");
            }

            using (GraphicsPath colorMatrixPath =
                BuildEnhancedSvgElementPath(
                    colorMatrixRect,
                    new RectangleF(
                        0f,
                        0f,
                        160f,
                        80f),
                    0f,
                    0f,
                    4f,
                    4f))
            using (Bitmap colorMatrixBitmap =
                new Bitmap(
                    160,
                    80,
                    PixelFormat.Format32bppArgb))
            using (Graphics colorMatrixGraphics =
                Graphics.FromImage(
                    colorMatrixBitmap))
            {
                colorMatrixGraphics.Clear(
                    Color.White);

                if (!DrawSvgColorMatrixApproximation(
                        colorMatrixGraphics,
                        colorMatrixRect,
                        document,
                        colorMatrixPath,
                        new RectangleF(
                            0f,
                            0f,
                            160f,
                            80f),
                        4f,
                        4f))
                {
                    throw new InvalidOperationException(
                        "SVG feColorMatrix approximation did not render.");
                }

                Color filteredPixel =
                    colorMatrixBitmap.GetPixel(
                        72,
                        20);

                if (filteredPixel.B <
                        150 ||
                    filteredPixel.R >
                        110)
                {
                    throw new InvalidOperationException(
                        "SVG feColorMatrix rendered output did not use the transformed color.");
                }
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

            string inheritedColorValue =
                GetSvgStyleInherited(
                    useTriangle,
                    "color");

            if (string.IsNullOrEmpty(
                    inheritedColorValue) ||
                inheritedColorValue.IndexOf(
                    "hsl(",
                    StringComparison.OrdinalIgnoreCase) < 0)
            {
                throw new InvalidOperationException(
                    "SVG inherit keyword did not continue to the parent style.");
            }

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
