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
        public static void ValidateChartAndSmartArtRendering(
            string outputDirectory)
        {
            if (string.IsNullOrEmpty(outputDirectory))
                throw new ArgumentException(
                    "PPTX fidelity output directory is required.",
                    "outputDirectory");

            outputDirectory = Path.GetFullPath(outputDirectory);
            Directory.CreateDirectory(outputDirectory);

            ValidateSyntheticChart(
                Path.Combine(
                    outputDirectory,
                    "chart-negative-values.png"));

            ValidateSyntheticSmartArt(
                Path.Combine(
                    outputDirectory,
                    "smartart-layouts.png"));
        }

        private static void ValidateSyntheticChart(
            string outputPath)
        {
            XmlDocument chart =
                new XmlDocument();

            chart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart>" +
                "<c:title><c:tx><c:rich><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Quarterly Delta</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
                "<c:plotArea><c:barChart><c:barDir val=\"col\"/>" +
                "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>" +
                BuildSyntheticChartSeries(
                    0,
                    "North",
                    new string[] { "Q1", "Q2", "Q3", "Q4" },
                    new double[] { 42.0, -26.0, 68.0, -12.0 }) +
                BuildSyntheticChartSeries(
                    1,
                    "South",
                    new string[] { "Q1", "Q2", "Q3", "Q4" },
                    new double[] { -18.0, 34.0, 51.0, -39.0 }) +
                "</c:barChart></c:plotArea>" +
                "</c:chart></c:chartSpace>");

            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            theme["accent1"] =
                Color.FromArgb(74, 122, 206);
            theme["accent2"] =
                Color.FromArgb(210, 86, 72);

            using (Bitmap bitmap =
                new Bitmap(
                    1000,
                    560,
                    PixelFormat.Format32bppArgb))
            using (Graphics graphics =
                Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode =
                    System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                DrawChart(
                    graphics,
                    chart,
                    new RectangleF(
                        20f,
                        20f,
                        960f,
                        520f),
                    theme);

                int nonWhite;
                int distinct;
                AnalyzeDiagnosticBitmap(
                    bitmap,
                    out nonWhite,
                    out distinct);

                if (nonWhite < 250 ||
                    distinct < 8)
                {
                    throw new InvalidOperationException(
                        "Synthetic chart rendering did not produce enough visual detail.");
                }

                bitmap.Save(
                    outputPath,
                    ImageFormat.Png);
            }

            RequireDiagnosticFile(outputPath);
        }

        private static string BuildSyntheticChartSeries(
            int index,
            string name,
            string[] categories,
            double[] values)
        {
            System.Text.StringBuilder xml =
                new System.Text.StringBuilder();

            xml.Append("<c:ser>");
            xml.Append("<c:idx val=\"");
            xml.Append(index.ToString());
            xml.Append("\"/><c:order val=\"");
            xml.Append(index.ToString());
            xml.Append("\"/>");
            xml.Append("<c:tx><c:v>");
            xml.Append(
                System.Security.SecurityElement.Escape(
                    name ?? string.Empty));
            xml.Append("</c:v></c:tx>");

            xml.Append("<c:cat><c:strLit>");
            for (int i = 0;
                 i < categories.Length;
                 i++)
            {
                xml.Append("<c:pt idx=\"");
                xml.Append(i.ToString());
                xml.Append("\"><c:v>");
                xml.Append(
                    System.Security.SecurityElement.Escape(
                        categories[i] ?? string.Empty));
                xml.Append("</c:v></c:pt>");
            }
            xml.Append("</c:strLit></c:cat>");

            xml.Append("<c:val><c:numLit>");
            for (int i = 0;
                 i < values.Length;
                 i++)
            {
                xml.Append("<c:pt idx=\"");
                xml.Append(i.ToString());
                xml.Append("\"><c:v>");
                xml.Append(
                    values[i].ToString(
                        "0.###",
                        System.Globalization.CultureInfo.InvariantCulture));
                xml.Append("</c:v></c:pt>");
            }
            xml.Append("</c:numLit></c:val>");
            xml.Append("</c:ser>");

            return xml.ToString();
        }

        private static void ValidateSyntheticSmartArt(
            string outputPath)
        {
            XmlDocument data =
                new XmlDocument();

            data.LoadXml(
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">" +
                "<dgm:ptLst>" +
                SyntheticSmartPoint("1", "Plan") +
                SyntheticSmartPoint("2", "Design") +
                SyntheticSmartPoint("3", "Build") +
                SyntheticSmartPoint("4", "Test") +
                SyntheticSmartPoint("5", "Release") +
                "</dgm:ptLst>" +
                "<dgm:cxnLst>" +
                SyntheticSmartConnection("1", "2") +
                SyntheticSmartConnection("2", "3") +
                SyntheticSmartConnection("3", "4") +
                SyntheticSmartConnection("4", "5") +
                "</dgm:cxnLst>" +
                "</dgm:dataModel>");

            string[] kinds =
                new string[]
                {
                    "Basic Process",
                    "Cycle",
                    "Matrix",
                    "Pyramid"
                };

            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            theme["accent1"] =
                Color.FromArgb(76, 123, 205);
            theme["accent2"] =
                Color.FromArgb(216, 102, 74);
            theme["accent3"] =
                Color.FromArgb(91, 157, 92);
            theme["accent4"] =
                Color.FromArgb(143, 106, 179);

            using (Bitmap bitmap =
                new Bitmap(
                    1200,
                    760,
                    PixelFormat.Format32bppArgb))
            using (Graphics graphics =
                Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode =
                    System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                for (int i = 0;
                     i < kinds.Length;
                     i++)
                {
                    XmlDocument layout =
                        new XmlDocument();

                    layout.LoadXml(
                        "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">" +
                        "<dgm:title val=\"" +
                        kinds[i] +
                        "\"/>" +
                        "</dgm:layoutDef>");

                    int column =
                        i % 2;
                    int row =
                        i / 2;

                    RectangleF rect =
                        new RectangleF(
                            20f +
                                column *
                                590f,
                            20f +
                                row *
                                365f,
                            570f,
                            345f);

                    using (Pen border =
                        new Pen(
                            Color.FromArgb(
                                225,
                                228,
                                232),
                            1f))
                    {
                        graphics.DrawRectangle(
                            border,
                            rect.X,
                            rect.Y,
                            rect.Width,
                            rect.Height);
                    }

                    if (!DrawStructuredSmartArt(
                            graphics,
                            data,
                            layout,
                            new RectangleF(
                                rect.X + 8f,
                                rect.Y + 8f,
                                rect.Width - 16f,
                                rect.Height - 16f),
                            theme))
                    {
                        throw new InvalidOperationException(
                            "Synthetic SmartArt layout failed: " +
                            kinds[i]);
                    }
                }

                int nonWhite;
                int distinct;
                AnalyzeDiagnosticBitmap(
                    bitmap,
                    out nonWhite,
                    out distinct);

                if (nonWhite < 350 ||
                    distinct < 7)
                {
                    throw new InvalidOperationException(
                        "Synthetic SmartArt rendering did not produce enough visual detail.");
                }

                bitmap.Save(
                    outputPath,
                    ImageFormat.Png);
            }

            RequireDiagnosticFile(outputPath);
        }

        private static string SyntheticSmartPoint(
            string id,
            string label)
        {
            return
                "<dgm:pt modelId=\"" +
                System.Security.SecurityElement.Escape(id) +
                "\"><dgm:t><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>" +
                System.Security.SecurityElement.Escape(label) +
                "</a:t></a:r></a:p></dgm:t></dgm:pt>";
        }

        private static string SyntheticSmartConnection(
            string source,
            string destination)
        {
            return
                "<dgm:cxn srcId=\"" +
                System.Security.SecurityElement.Escape(source) +
                "\" destId=\"" +
                System.Security.SecurityElement.Escape(destination) +
                "\"/>";
        }

        private static void AnalyzeDiagnosticBitmap(
            Bitmap bitmap,
            out int nonWhite,
            out int distinct)
        {
            nonWhite = 0;
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

                    nonWhite++;

                    int bucket =
                        ((color.R / 32) << 10) |
                        ((color.G / 32) << 5) |
                        (color.B / 32);

                    colors.Add(bucket);
                }
            }

            distinct = colors.Count;
        }

        private static void RequireDiagnosticFile(
            string path)
        {
            FileInfo info =
                new FileInfo(path);

            if (!info.Exists ||
                info.Length <= 0)
            {
                throw new InvalidOperationException(
                    "Diagnostic PNG was not created: " +
                    path);
            }
        }
    }
}
