using System;
using System.Collections.Generic;
using System.Drawing;
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
                "<linearGradient id=\"g\" x1=\"0%\" y1=\"0%\" x2=\"100%\" y2=\"0%\" gradientTransform=\"rotate(22 .5 .5)\">" +
                "<stop offset=\"0%\" stop-color=\"#e84b4b\"/>" +
                "<stop offset=\"50%\" stop-color=\"#f2c94c\"/>" +
                "<stop offset=\"100%\" stop-color=\"#4b72e8\"/>" +
                "</linearGradient>" +
                "<clipPath id=\"clip\"><circle cx=\"100\" cy=\"78\" r=\"40\"/></clipPath>" +
                "</defs>" +
                "<g transform=\"matrix(1 0.10 -0.08 1 3 1)\"><rect x=\"16\" y=\"12\" width=\"58\" height=\"28\" rx=\"6\" fill=\"#20a77a\"/></g>" +
                "<circle cx=\"154\" cy=\"28\" r=\"18\" fill=\"url(#g)\" transform=\"skewX(8)\"/>" +
                "<path d=\"M 12 52 C 30 38 42 68 60 52 S 90 38 108 52 Q 126 70 142 52 T 184 52\" fill=\"none\" stroke=\"#6f42a8\" stroke-width=\"2\"/>" +
                "<path d=\"M 18 88 A 82 48 0 0 1 182 88 L 182 116 L 18 116 Z\" " +
                "fill=\"url(#g)\" clip-path=\"url(#clip)\" stroke=\"#243447\" stroke-width=\"1.5\"/>" +
                "<text x=\"100\" y=\"63\" font-size=\"14\" fill=\"#20252b\">SVG</text>" +
                "</svg>";

            XmlDocument document =
                new XmlDocument();
            document.LoadXml(svgText);

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
