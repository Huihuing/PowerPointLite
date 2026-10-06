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
        public static void ValidatePptxFidelityRendering(
            string outputDirectory)
        {
            ValidateChartAndSmartArtRendering(
                outputDirectory);

            ValidateSyntheticRichText(
                Path.Combine(
                    Path.GetFullPath(outputDirectory),
                    "rich-text-layout.png"));
        }

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
                "<c:dLbls><c:showVal val=\"1\"/><c:showCatName val=\"1\"/><c:separator> · </c:separator></c:dLbls>" +
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
                "</c:barChart>" +
                "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/><c:min val=\"-50\"/><c:max val=\"80\"/></c:scaling><c:majorUnit val=\"20\"/></c:valAx>" +
                "</c:plotArea>" +
                "<c:legend><c:legendPos val=\"b\"/></c:legend>" +
                "</c:chart></c:chartSpace>");

            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            theme["accent1"] =
                Color.FromArgb(74, 122, 206);
            theme["accent2"] =
                Color.FromArgb(210, 86, 72);

            List<XmlNode> syntheticSeries =
                FindAll(
                    chart,
                    "ser");

            if (syntheticSeries.Count < 2)
            {
                throw new InvalidOperationException(
                    "Synthetic chart series were not created.");
            }

            Color? firstSeriesColor =
                ReadChartSeriesColor(
                    syntheticSeries[0],
                    theme);

            if (!firstSeriesColor.HasValue ||
                firstSeriesColor.Value.R != 0x33 ||
                firstSeriesColor.Value.G != 0x66 ||
                firstSeriesColor.Value.B != 0xCC)
            {
                throw new InvalidOperationException(
                    "Explicit chart series color was not parsed.");
            }

            ChartAxisScale scale =
                ReadChartAxisScale(
                    chart,
                    -39.0,
                    68.0);

            if (Math.Abs(
                    scale.Minimum -
                    (-50.0)) > 0.0001 ||
                Math.Abs(
                    scale.Maximum -
                    80.0) > 0.0001 ||
                Math.Abs(
                    scale.MajorUnit -
                    20.0) > 0.0001 ||
                scale.Reverse)
            {
                throw new InvalidOperationException(
                    "Explicit chart value-axis scaling was not retained.");
            }

            ChartLabelOptions labels =
                ReadChartLabelOptions(
                    chart);

            if (!labels.ShowValue ||
                !labels.ShowCategoryName ||
                labels.ShowSeriesName ||
                labels.ShowPercent ||
                labels.Separator != " · ")
            {
                throw new InvalidOperationException(
                    "Chart data-label options were not parsed correctly.");
            }

            if (ReadChartLegendPosition(chart) != "b")
            {
                throw new InvalidOperationException(
                    "Chart legend position was not parsed correctly.");
            }

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

        private static void ValidateSyntheticRichText(
            string outputPath)
        {
            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            using (Bitmap bitmap =
                new Bitmap(
                    1100,
                    640,
                    PixelFormat.Format32bppArgb))
            using (Graphics graphics =
                Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode =
                    System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                DrawDiagnosticTextFrame(
                    graphics,
                    new RectangleF(
                        20f,
                        20f,
                        650f,
                        100f),
                    BuildSyntheticTabbedShape(),
                    theme);

                DrawDiagnosticTextFrame(
                    graphics,
                    new RectangleF(
                        20f,
                        145f,
                        650f,
                        85f),
                    BuildSyntheticRtlShape(),
                    theme);

                RectangleF autoFitRect =
                    new RectangleF(
                        20f,
                        270f,
                        650f,
                        72f);

                DrawDiagnosticTextFrame(
                    graphics,
                    autoFitRect,
                    BuildSyntheticShapeAutoFit(),
                    theme);

                DrawDiagnosticTextFrame(
                    graphics,
                    new RectangleF(
                        710f,
                        20f,
                        360f,
                        210f),
                    BuildSyntheticNormalAutoFit(),
                    theme);

                int nonWhite;
                int distinct;
                AnalyzeDiagnosticBitmap(
                    bitmap,
                    out nonWhite,
                    out distinct);

                if (nonWhite < 220 ||
                    distinct < 3)
                {
                    throw new InvalidOperationException(
                        "Synthetic rich-text rendering did not produce enough visual detail.");
                }

                int belowAutoFit =
                    CountNonWhiteDiagnosticPixels(
                        bitmap,
                        new Rectangle(
                            20,
                            (int)Math.Ceiling(
                                autoFitRect.Bottom + 2f),
                            650,
                            155));

                if (belowAutoFit < 4)
                {
                    throw new InvalidOperationException(
                        "spAutoFit approximation did not extend overflowing text beyond the original text rectangle.");
                }

                int tabRightSide =
                    CountNonWhiteDiagnosticPixels(
                        bitmap,
                        new Rectangle(
                            210,
                            20,
                            460,
                            100));

                if (tabRightSide < 4)
                {
                    throw new InvalidOperationException(
                        "Tab-stop rendering did not place later text runs across the line.");
                }

                bitmap.Save(
                    outputPath,
                    ImageFormat.Png);
            }

            RequireDiagnosticFile(outputPath);
        }

        private static void DrawDiagnosticTextFrame(
            Graphics graphics,
            RectangleF rect,
            XmlDocument shapeDocument,
            Dictionary<string, Color> theme)
        {
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

            DrawRichShapeText(
                graphics,
                shapeDocument.DocumentElement,
                rect,
                theme,
                7);
        }

        private static XmlDocument BuildSyntheticTabbedShape()
        {
            return LoadSyntheticTextShape(
                "<a:bodyPr lIns=\"45720\" rIns=\"45720\" tIns=\"22860\" bIns=\"22860\" defTabSz=\"914400\"/>" +
                "<a:p><a:pPr><a:tabLst>" +
                "<a:tab pos=\"1371600\"/><a:tab pos=\"2743200\"/>" +
                "</a:tabLst></a:pPr>" +
                "<a:r><a:rPr sz=\"1800\" b=\"1\"/><a:t>Alpha</a:t></a:r>" +
                "<a:tab/>" +
                "<a:r><a:rPr sz=\"1800\"/><a:t>Beta</a:t></a:r>" +
                "<a:tab/>" +
                "<a:r><a:rPr sz=\"1800\"/><a:t>Gamma</a:t></a:r>" +
                "</a:p>");
        }

        private static XmlDocument BuildSyntheticRtlShape()
        {
            return LoadSyntheticTextShape(
                "<a:bodyPr lIns=\"45720\" rIns=\"45720\" tIns=\"22860\" bIns=\"22860\"/>" +
                "<a:p><a:pPr rtl=\"1\" algn=\"r\"/>" +
                "<a:r><a:rPr sz=\"1900\"/><a:t>RTL paragraph 123</a:t></a:r>" +
                "</a:p>");
        }

        private static XmlDocument BuildSyntheticShapeAutoFit()
        {
            return LoadSyntheticTextShape(
                "<a:bodyPr lIns=\"45720\" rIns=\"45720\" tIns=\"22860\" bIns=\"22860\"><a:spAutoFit/></a:bodyPr>" +
                "<a:p><a:r><a:rPr sz=\"1700\"/><a:t>Auto fit line one</a:t></a:r>" +
                "<a:br/><a:r><a:rPr sz=\"1700\"/><a:t>Auto fit line two</a:t></a:r>" +
                "<a:br/><a:r><a:rPr sz=\"1700\"/><a:t>Auto fit line three</a:t></a:r>" +
                "<a:br/><a:r><a:rPr sz=\"1700\"/><a:t>Auto fit line four</a:t></a:r>" +
                "</a:p>");
        }

        private static XmlDocument BuildSyntheticNormalAutoFit()
        {
            return LoadSyntheticTextShape(
                "<a:bodyPr lIns=\"45720\" rIns=\"45720\" tIns=\"22860\" bIns=\"22860\" wrap=\"square\">" +
                "<a:normAutofit fontScale=\"92000\" lnSpcReduction=\"10000\"/></a:bodyPr>" +
                "<a:p><a:pPr algn=\"ctr\" spcBef=\"0\"/>" +
                "<a:r><a:rPr sz=\"2400\" b=\"1\"/><a:t>This deliberately long text should wrap and shrink to stay inside the smaller box.</a:t></a:r>" +
                "</a:p>");
        }

        private static XmlDocument LoadSyntheticTextShape(
            string textBodyContent)
        {
            XmlDocument document =
                new XmlDocument();

            document.LoadXml(
                "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<p:txBody>" +
                textBodyContent +
                "</p:txBody>" +
                "</p:sp>");

            return document;
        }

        private static int CountNonWhiteDiagnosticPixels(
            Bitmap bitmap,
            Rectangle area)
        {
            if (bitmap == null)
                return 0;

            Rectangle bounds =
                Rectangle.Intersect(
                    new Rectangle(
                        0,
                        0,
                        bitmap.Width,
                        bitmap.Height),
                    area);

            int count = 0;

            for (int y = bounds.Top;
                 y < bounds.Bottom;
                 y += 4)
            {
                for (int x = bounds.Left;
                     x < bounds.Right;
                     x += 4)
                {
                    Color color =
                        bitmap.GetPixel(
                            x,
                            y);

                    if (color.R < 245 ||
                        color.G < 245 ||
                        color.B < 245)
                    {
                        count++;
                    }
                }
            }

            return count;
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

            string color =
                index == 0
                    ? "3366CC"
                    : "CC5533";

            xml.Append(
                "<c:spPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:srgbClr val=\"");
            xml.Append(color);
            xml.Append(
                "\"/></a:solidFill></c:spPr>");

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
