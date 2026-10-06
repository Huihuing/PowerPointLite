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
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 200 120\">" +
                "<defs>" +
                "<linearGradient id=\"g\" x1=\"0%\" y1=\"0%\" x2=\"100%\" y2=\"0%\" spreadMethod=\"reflect\" gradientTransform=\"rotate(22 .5 .5)\">" +
                "<stop offset=\"0%\" stop-color=\"#e84b4b\"/>" +
                "<stop offset=\"50%\" stop-color=\"#f2c94c\"/>" +
                "<stop offset=\"100%\" stop-color=\"#4b72e8\"/>" +
                "</linearGradient>" +
                "<clipPath id=\"clip\"><circle cx=\"100\" cy=\"78\" r=\"40\"/></clipPath>" +
                "<clipPath id=\"nestedClip\"><g transform=\"translate(22 0)\"><rect x=\"10\" y=\"8\" width=\"18\" height=\"16\"/></g></clipPath>" +
                "<pattern id=\"pat\" patternUnits=\"userSpaceOnUse\" width=\"12\" height=\"12\">" +
                "<rect x=\"0\" y=\"0\" width=\"6\" height=\"12\" fill=\"#e85d75\"/>" +
                "<rect x=\"6\" y=\"0\" width=\"6\" height=\"12\" fill=\"#4c78d6\"/>" +
                "</pattern>" +
                "<mask id=\"mask\"><circle cx=\"158\" cy=\"88\" r=\"24\" fill=\"white\"/></mask>" +
                "</defs>" +
                "<g transform=\"matrix(1 0.10 -0.08 1 3 1)\"><rect x=\"16\" y=\"12\" width=\"58\" height=\"28\" rx=\"6\" fill=\"#20a77a\"/></g>" +
                "<circle cx=\"154\" cy=\"28\" r=\"18\" fill=\"url(#g)\" transform=\"skewX(8)\"/>" +
                "<rect x=\"8\" y=\"6\" width=\"48\" height=\"22\" fill=\"#1677d2\" clip-path=\"url(#nestedClip)\"/>" +
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
