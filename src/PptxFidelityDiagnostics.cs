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
        public static void ValidatePptxFidelityRendering(
            string outputDirectory)
        {
            ValidateChartAndSmartArtRendering(
                outputDirectory);

            ValidateSyntheticRichText(
                Path.Combine(
                    Path.GetFullPath(outputDirectory),
                    "rich-text-layout.png"));

            ValidateSyntheticCustomGeometry(
                Path.Combine(
                    Path.GetFullPath(outputDirectory),
                    "custom-geometry.png"));

            ValidateSyntheticImageEffects(
                Path.Combine(
                    Path.GetFullPath(outputDirectory),
                    "image-effects.png"));

            ValidateSyntheticShapeEffects(
                Path.Combine(
                    Path.GetFullPath(outputDirectory),
                    "shape-effects.png"));

            ValidateSyntheticColorTransforms();

            ValidateSyntheticTextInheritance();
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

            ValidateSyntheticExtendedCharts(
                Path.Combine(
                    outputDirectory,
                    "extended-charts.png"));
        }

        private static void ValidateSyntheticChart(
            string outputPath)
        {
            XmlDocument chart =
                new XmlDocument();

            chart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:spPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:srgbClr val=\"F5F7FB\"/></a:solidFill></c:spPr>" +
                "<c:chart>" +
                "<c:title><c:tx><c:rich><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Quarterly Delta</a:t></a:r></a:p></c:rich></c:tx><c:txPr><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"1500\" b=\"1\" i=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"228844\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:title>" +
                "<c:plotArea><c:barChart><c:barDir val=\"col\"/>" +
                "<c:dLbls><c:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/><c:dLblPos val=\"inEnd\"/><c:showVal val=\"1\"/><c:showCatName val=\"1\"/><c:spPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:srgbClr val=\"FFF1CC\"/></a:solidFill><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"12700\"><a:solidFill><a:srgbClr val=\"775511\"/></a:solidFill><a:prstDash val=\"dot\"/></a:ln></c:spPr><c:txPr><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"1350\" b=\"1\" i=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"224488\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:separator> · </c:separator></c:dLbls>" +
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
                "<c:catAx><c:axId val=\"1\"/><c:tickLblPos val=\"none\"/><c:tickLblSkip val=\"2\"/><c:majorTickMark val=\"out\"/><c:crossesAt val=\"20\"/><c:spPr><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"25400\"><a:solidFill><a:srgbClr val=\"3366CC\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></c:spPr><c:txPr><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"900\" b=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"2255AA\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:title><c:tx><c:rich><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx><c:txPr><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"1050\" b=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"7733AA\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:title></c:catAx>" +
                "<c:valAx><c:axId val=\"2\"/><c:tickLblPos val=\"none\"/><c:majorTickMark val=\"cross\"/><c:minorTickMark val=\"in\"/><c:scaling><c:orientation val=\"minMax\"/><c:min val=\"-50\"/><c:max val=\"80\"/></c:scaling><c:majorUnit val=\"20\"/><c:minorUnit val=\"10\"/><c:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/><c:majorGridlines><c:spPr><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"12700\"><a:solidFill><a:srgbClr val=\"88AACC\"/></a:solidFill><a:prstDash val=\"dashDot\"/></a:ln></c:spPr></c:majorGridlines><c:minorGridlines><c:spPr><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"9525\"><a:solidFill><a:srgbClr val=\"CCDDEE\"/></a:solidFill><a:prstDash val=\"dot\"/></a:ln></c:spPr></c:minorGridlines><c:spPr><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"19050\"><a:solidFill><a:srgbClr val=\"CC5533\"/></a:solidFill><a:prstDash val=\"dot\"/></a:ln></c:spPr><c:txPr><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"1000\" i=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"AA5522\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:title><c:tx><c:rich><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Delta</a:t></a:r></a:p></c:rich></c:tx><c:txPr><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"1150\" i=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"CC7722\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:title></c:valAx>" +
                "<c:spPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:srgbClr val=\"FFF8EE\"/></a:solidFill></c:spPr>" +
                "</c:plotArea>" +
                "<c:legend><c:legendPos val=\"b\"/><c:legendEntry><c:idx val=\"1\"/><c:delete val=\"1\"/></c:legendEntry><c:txPr><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"1200\" b=\"1\" i=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"CC2277\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:legend>" +
                "</c:chart></c:chartSpace>");

            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            theme["accent1"] =
                Color.FromArgb(74, 122, 206);
            theme["accent2"] =
                Color.FromArgb(210, 86, 72);

            Color chartAreaFill =
                ReadChartAreaFill(
                    chart,
                    theme,
                    "chartSpace",
                    Color.Black);
            Color plotAreaFill =
                ReadChartAreaFill(
                    chart,
                    theme,
                    "plotArea",
                    Color.Black);

            if (chartAreaFill.R != 0xF5 ||
                chartAreaFill.G != 0xF7 ||
                chartAreaFill.B != 0xFB ||
                plotAreaFill.R != 0xFF ||
                plotAreaFill.G != 0xF8 ||
                plotAreaFill.B != 0xEE)
            {
                throw new InvalidOperationException(
                    "ChartSpace or plotArea solid fill was not parsed correctly.");
            }

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

            ChartSeriesData visual =
                new ChartSeriesData();

            ReadChartSeriesVisualStyle(
                syntheticSeries[0],
                visual);

            if (visual.LineDashStyle !=
                    System.Drawing.Drawing2D.DashStyle.Dash ||
                visual.MarkerSymbol != "diamond" ||
                Math.Abs(
                    visual.MarkerSize -
                    9f) > 0.01f ||
                !visual.MarkerEnabled)
            {
                throw new InvalidOperationException(
                    "Chart marker/line visual style was not parsed.");
            }

            List<ChartSeriesData> extendedSeries =
                ReadStandardChartSeries(
                    chart,
                    theme);

            if (extendedSeries.Count < 2 ||
                !extendedSeries[0].ExplicitColor.HasValue ||
                extendedSeries[0].ExplicitColor.Value.R != 0x33 ||
                extendedSeries[0].ExplicitColor.Value.G != 0x66 ||
                extendedSeries[0].ExplicitColor.Value.B != 0xCC ||
                extendedSeries[0].LineDashStyle !=
                    System.Drawing.Drawing2D.DashStyle.Dash ||
                extendedSeries[0].MarkerSymbol != "diamond" ||
                Math.Abs(
                    extendedSeries[0].MarkerSize -
                    9f) > 0.01f)
            {
                throw new InvalidOperationException(
                    "Extended chart series parser did not preserve common series styling.");
            }

            if (ReadChartAxisTitle(
                    chart,
                    "catAx") != "Quarter" ||
                ReadChartAxisTitle(
                    chart,
                    "valAx") != "Delta")
            {
                throw new InvalidOperationException(
                    "Chart axis titles were not parsed.");
            }

            ChartLabelOptions categoryAxisTitleStyle =
                ReadChartAxisTitleTextStyle(
                    chart,
                    "catAx",
                    theme);
            ChartLabelOptions valueAxisTitleStyle =
                ReadChartAxisTitleTextStyle(
                    chart,
                    "valAx",
                    theme);

            if (!categoryAxisTitleStyle.TextColor.HasValue ||
                categoryAxisTitleStyle.TextColor.Value.R != 0x77 ||
                categoryAxisTitleStyle.TextColor.Value.G != 0x33 ||
                categoryAxisTitleStyle.TextColor.Value.B != 0xAA ||
                !categoryAxisTitleStyle.FontSize.HasValue ||
                Math.Abs(
                    categoryAxisTitleStyle.FontSize.Value -
                    10.5f) > 0.01f ||
                !categoryAxisTitleStyle.Bold.HasValue ||
                !categoryAxisTitleStyle.Bold.Value ||
                categoryAxisTitleStyle.FontFamily != "Arial" ||
                !valueAxisTitleStyle.TextColor.HasValue ||
                valueAxisTitleStyle.TextColor.Value.R != 0xCC ||
                valueAxisTitleStyle.TextColor.Value.G != 0x77 ||
                valueAxisTitleStyle.TextColor.Value.B != 0x22 ||
                !valueAxisTitleStyle.FontSize.HasValue ||
                Math.Abs(
                    valueAxisTitleStyle.FontSize.Value -
                    11.5f) > 0.01f ||
                !valueAxisTitleStyle.Italic.HasValue ||
                !valueAxisTitleStyle.Italic.Value ||
                valueAxisTitleStyle.FontFamily != "Arial")
            {
                throw new InvalidOperationException(
                    "Chart axis title txPr text style was not parsed correctly.");
            }

            ChartLabelOptions chartTitleStyle =
                ReadChartTitleTextStyle(
                    chart,
                    theme);

            if (!chartTitleStyle.TextColor.HasValue ||
                chartTitleStyle.TextColor.Value.R != 0x22 ||
                chartTitleStyle.TextColor.Value.G != 0x88 ||
                chartTitleStyle.TextColor.Value.B != 0x44 ||
                !chartTitleStyle.FontSize.HasValue ||
                Math.Abs(
                    chartTitleStyle.FontSize.Value -
                    15f) > 0.01f ||
                !chartTitleStyle.Bold.HasValue ||
                !chartTitleStyle.Bold.Value ||
                !chartTitleStyle.Italic.HasValue ||
                !chartTitleStyle.Italic.Value ||
                chartTitleStyle.FontFamily != "Arial")
            {
                throw new InvalidOperationException(
                    "Chart title txPr text style was not parsed correctly.");
            }

            if (!string.Equals(
                    ReadChartAxisTickLabelPosition(
                        chart,
                        "catAx"),
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Chart category-axis tick label visibility was not parsed.");
            }

            if (!string.Equals(
                    ReadChartAxisTickLabelPosition(
                        chart,
                        "valAx"),
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Chart value-axis tick label visibility was not parsed.");
            }

            if (ReadChartCategoryLabelSkip(
                    chart) != 2)
            {
                throw new InvalidOperationException(
                    "Chart category-axis tickLblSkip was not parsed.");
            }

            if (ReadChartAxisTickMark(
                    chart,
                    "catAx",
                    "majorTickMark") != "out" ||
                ReadChartAxisTickMark(
                    chart,
                    "valAx",
                    "majorTickMark") != "cross" ||
                ReadChartAxisTickMark(
                    chart,
                    "valAx",
                    "minorTickMark") != "in")
            {
                throw new InvalidOperationException(
                    "Chart major/minor tick mark settings were not parsed.");
            }

            ChartLabelOptions categoryAxisTextStyle =
                ReadChartAxisLabelStyle(
                    chart,
                    "catAx",
                    theme);
            ChartLabelOptions valueAxisTextStyle =
                ReadChartAxisLabelStyle(
                    chart,
                    "valAx",
                    theme);

            if (!categoryAxisTextStyle.TextColor.HasValue ||
                categoryAxisTextStyle.TextColor.Value.R != 0x22 ||
                categoryAxisTextStyle.TextColor.Value.G != 0x55 ||
                categoryAxisTextStyle.TextColor.Value.B != 0xAA ||
                !categoryAxisTextStyle.FontSize.HasValue ||
                Math.Abs(
                    categoryAxisTextStyle.FontSize.Value -
                    9f) > 0.01f ||
                !categoryAxisTextStyle.Bold.HasValue ||
                !categoryAxisTextStyle.Bold.Value ||
                categoryAxisTextStyle.FontFamily != "Arial" ||
                !valueAxisTextStyle.TextColor.HasValue ||
                valueAxisTextStyle.TextColor.Value.R != 0xAA ||
                valueAxisTextStyle.TextColor.Value.G != 0x55 ||
                valueAxisTextStyle.TextColor.Value.B != 0x22 ||
                !valueAxisTextStyle.FontSize.HasValue ||
                Math.Abs(
                    valueAxisTextStyle.FontSize.Value -
                    10f) > 0.01f ||
                !valueAxisTextStyle.Italic.HasValue ||
                !valueAxisTextStyle.Italic.Value ||
                valueAxisTextStyle.FontFamily != "Arial")
            {
                throw new InvalidOperationException(
                    "Chart axis txPr text style was not parsed correctly.");
            }

            using (Bitmap axisTitleBitmap =
                new Bitmap(
                    360,
                    260,
                    PixelFormat.Format32bppArgb))
            using (Graphics axisTitleGraphics =
                Graphics.FromImage(
                    axisTitleBitmap))
            {
                axisTitleGraphics.Clear(
                    Color.White);

                DrawChartAxisTitles(
                    axisTitleGraphics,
                    new RectangleF(
                        70f,
                        30f,
                        250f,
                        150f),
                    new RectangleF(
                        20f,
                        10f,
                        320f,
                        230f),
                    "column",
                    "Delta",
                    "Quarter",
                    valueAxisTitleStyle,
                    categoryAxisTitleStyle);

                bool categoryTitleVisible =
                    false;
                bool valueTitleVisible =
                    false;

                for (int y = 190;
                     y < 238 &&
                     !categoryTitleVisible;
                     y++)
                {
                    for (int x = 90;
                         x < 305;
                         x++)
                    {
                        Color pixel =
                            axisTitleBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.B >
                                pixel.R + 25 &&
                            pixel.R >
                                pixel.G + 35 &&
                            pixel.B > 100)
                        {
                            categoryTitleVisible =
                                true;
                            break;
                        }
                    }
                }

                for (int y = 55;
                     y < 180 &&
                     !valueTitleVisible;
                     y++)
                {
                    for (int x = 15;
                         x < 38;
                         x++)
                    {
                        Color pixel =
                            axisTitleBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R >
                                pixel.G + 50 &&
                            pixel.R >
                                pixel.B + 80 &&
                            pixel.R > 130)
                        {
                            valueTitleVisible =
                                true;
                            break;
                        }
                    }
                }

                if (!categoryTitleVisible ||
                    !valueTitleVisible)
                {
                    throw new InvalidOperationException(
                        "Chart axis title txPr colors were not rendered.");
                }
            }

            ChartLineStyle categoryAxisStyle =
                ReadChartAxisLineStyle(
                    chart,
                    "catAx",
                    theme);
            ChartLineStyle valueAxisStyle =
                ReadChartAxisLineStyle(
                    chart,
                    "valAx",
                    theme);

            if (categoryAxisStyle.Color.R != 0x33 ||
                categoryAxisStyle.Color.G != 0x66 ||
                categoryAxisStyle.Color.B != 0xCC ||
                categoryAxisStyle.Width <= 1f ||
                categoryAxisStyle.DashStyle !=
                    DashStyle.Dash ||
                valueAxisStyle.Color.R != 0xCC ||
                valueAxisStyle.Color.G != 0x55 ||
                valueAxisStyle.Color.B != 0x33 ||
                valueAxisStyle.Width <= 1f ||
                valueAxisStyle.DashStyle !=
                    DashStyle.Dot)
            {
                throw new InvalidOperationException(
                    "Chart axis line color, width, or dash style was not parsed.");
            }

            ChartLineStyle majorGridStyle =
                ReadChartMajorGridlineStyle(
                    chart,
                    theme);

            if (majorGridStyle.Color.R != 0x88 ||
                majorGridStyle.Color.G != 0xAA ||
                majorGridStyle.Color.B != 0xCC ||
                majorGridStyle.Width < 1f ||
                majorGridStyle.DashStyle !=
                    DashStyle.DashDot)
            {
                throw new InvalidOperationException(
                    "Chart major gridline color, width, or dash style was not parsed.");
            }

            ChartLineStyle minorGridStyle =
                ReadChartMinorGridlineStyle(
                    chart,
                    theme);

            if (minorGridStyle == null ||
                minorGridStyle.Color.R != 0xCC ||
                minorGridStyle.Color.G != 0xDD ||
                minorGridStyle.Color.B != 0xEE ||
                minorGridStyle.Width < 1f ||
                minorGridStyle.DashStyle !=
                    DashStyle.Dot)
            {
                throw new InvalidOperationException(
                    "Chart minor gridline color, width, or dash style was not parsed.");
            }

            ChartSeriesData highLabelSeries =
                new ChartSeriesData();
            highLabelSeries.Categories.Add(
                "HIGH");

            using (Bitmap labelBitmap =
                new Bitmap(
                    240,
                    140,
                    PixelFormat.Format32bppArgb))
            using (Graphics labelGraphics =
                Graphics.FromImage(
                    labelBitmap))
            {
                labelGraphics.Clear(
                    Color.White);

                RectangleF labelPlot =
                    new RectangleF(
                        60f,
                        55f,
                        120f,
                        45f);

                DrawChartCategoryLabels(
                    labelGraphics,
                    labelPlot,
                    highLabelSeries,
                    1,
                    "column",
                    "high",
                    1,
                    categoryAxisTextStyle);

                bool foundAbovePlot =
                    false;
                bool foundStyledCategoryColor =
                    false;

                for (int y = 20;
                     y < 55 &&
                     !foundAbovePlot;
                     y++)
                {
                    for (int x = 60;
                         x < 180;
                         x++)
                    {
                        Color pixel =
                            labelBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R < 220 ||
                            pixel.G < 220 ||
                            pixel.B < 220)
                        {
                            foundAbovePlot =
                                true;
                        }

                        if (pixel.B >
                                pixel.R + 45 &&
                            pixel.B >
                                pixel.G + 20)
                        {
                            foundStyledCategoryColor =
                                true;
                        }

                        if (foundAbovePlot &&
                            foundStyledCategoryColor)
                        {
                            break;
                        }
                    }
                }

                if (!foundAbovePlot ||
                    !foundStyledCategoryColor)
                {
                    throw new InvalidOperationException(
                        "Chart category tickLblPos=high or catAx txPr styling was not rendered.");
                }
            }

            ChartAxisScale highValueScale =
                new ChartAxisScale();
            highValueScale.Minimum =
                0.0;
            highValueScale.Maximum =
                10.0;
            highValueScale.MajorUnit =
                10.0;

            using (Bitmap valueLabelBitmap =
                new Bitmap(
                    260,
                    150,
                    PixelFormat.Format32bppArgb))
            using (Graphics valueLabelGraphics =
                Graphics.FromImage(
                    valueLabelBitmap))
            {
                valueLabelGraphics.Clear(
                    Color.White);

                RectangleF valuePlot =
                    new RectangleF(
                        60f,
                        35f,
                        120f,
                        80f);

                DrawChartValueGrid(
                    valueLabelGraphics,
                    valuePlot,
                    highValueScale,
                    "column",
                    true,
                    "high",
                    null,
                    null);

                bool foundRightOfPlot =
                    false;

                for (int y = 25;
                     y < 125 &&
                     !foundRightOfPlot;
                     y++)
                {
                    for (int x = 184;
                         x < 240;
                         x++)
                    {
                        Color pixel =
                            valueLabelBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R < 220 ||
                            pixel.G < 220 ||
                            pixel.B < 220)
                        {
                            foundRightOfPlot =
                                true;
                            break;
                        }
                    }
                }

                if (!foundRightOfPlot)
                {
                    throw new InvalidOperationException(
                        "Chart value tickLblPos=high did not move labels to the opposite side.");
                }
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
                Math.Abs(
                    scale.MinorUnit -
                    10.0) > 0.0001 ||
                scale.NumberFormat != "0.0" ||
                scale.Reverse)
            {
                throw new InvalidOperationException(
                    "Explicit chart value-axis scaling was not retained.");
            }

            if (Math.Abs(
                    ReadChartCategoryAxisCrossValue(
                        chart,
                        scale) -
                    20.0) > 0.0001)
            {
                throw new InvalidOperationException(
                    "Chart category-axis crossesAt value was not parsed.");
            }

            using (Bitmap tickBitmap =
                new Bitmap(
                    240,
                    160,
                    PixelFormat.Format32bppArgb))
            using (Graphics tickGraphics =
                Graphics.FromImage(
                    tickBitmap))
            {
                tickGraphics.Clear(
                    Color.White);

                RectangleF tickPlot =
                    new RectangleF(
                        50f,
                        20f,
                        140f,
                        100f);

                ChartLineStyle tickCategoryStyle =
                    new ChartLineStyle();
                tickCategoryStyle.Color =
                    Color.FromArgb(
                        30,
                        90,
                        220);
                tickCategoryStyle.Width =
                    2f;

                ChartLineStyle tickValueStyle =
                    new ChartLineStyle();
                tickValueStyle.Color =
                    Color.FromArgb(
                        220,
                        50,
                        40);
                tickValueStyle.Width =
                    2f;

                DrawChartAxisTickMarks(
                    tickGraphics,
                    tickPlot,
                    scale,
                    4,
                    "column",
                    0f,
                    80f,
                    chart,
                    tickCategoryStyle,
                    tickValueStyle);

                bool categoryOutVisible =
                    false;
                bool valueCrossVisible =
                    false;
                bool valueMinorInVisible =
                    false;

                for (int y = 81;
                     y <= 87;
                     y++)
                {
                    for (int x = 64;
                         x <= 71;
                         x++)
                    {
                        Color pixel =
                            tickBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.B >
                                pixel.R + 60 &&
                            pixel.B >
                                pixel.G + 40)
                        {
                            categoryOutVisible =
                                true;
                        }
                    }
                }

                for (int y = 78;
                     y <= 85;
                     y++)
                {
                    for (int x = 44;
                         x <= 56;
                         x++)
                    {
                        Color pixel =
                            tickBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R >
                                pixel.G + 70 &&
                            pixel.R >
                                pixel.B + 70)
                        {
                            valueCrossVisible =
                                true;
                        }
                    }
                }

                for (int y = 108;
                     y <= 116;
                     y++)
                {
                    for (int x = 51;
                         x <= 55;
                         x++)
                    {
                        Color pixel =
                            tickBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R >
                                pixel.G + 70 &&
                            pixel.R >
                                pixel.B + 70)
                        {
                            valueMinorInVisible =
                                true;
                        }
                    }
                }

                if (!categoryOutVisible ||
                    !valueCrossVisible ||
                    !valueMinorInVisible)
                {
                    throw new InvalidOperationException(
                        "Chart major/minor axis tick marks were not rendered in the requested direction.");
                }
            }

            List<double> diagnosticMajorTicks =
                BuildChartAxisTicks(
                    scale);
            List<double> diagnosticMinorTicks =
                BuildChartMinorAxisTicks(
                    scale,
                    diagnosticMajorTicks);

            if (diagnosticMinorTicks.Count == 0)
            {
                throw new InvalidOperationException(
                    "Chart minorUnit did not produce minor axis ticks.");
            }

            ChartLabelOptions labels =
                ReadChartLabelOptions(
                    chart,
                    theme);

            if (!labels.ShowValue ||
                !labels.ShowCategoryName ||
                labels.ShowSeriesName ||
                labels.ShowPercent ||
                !labels.TextColor.HasValue ||
                labels.TextColor.Value.R != 0x22 ||
                labels.TextColor.Value.G != 0x44 ||
                labels.TextColor.Value.B != 0x88 ||
                !labels.FillColor.HasValue ||
                labels.FillColor.Value.R != 0xFF ||
                labels.FillColor.Value.G != 0xF1 ||
                labels.FillColor.Value.B != 0xCC ||
                labels.BorderStyle == null ||
                labels.BorderStyle.Color.R != 0x77 ||
                labels.BorderStyle.Color.G != 0x55 ||
                labels.BorderStyle.Color.B != 0x11 ||
                labels.BorderStyle.DashStyle !=
                    DashStyle.Dot ||
                !labels.FontSize.HasValue ||
                Math.Abs(
                    labels.FontSize.Value -
                    13.5f) > 0.01f ||
                !labels.Bold.HasValue ||
                !labels.Bold.Value ||
                !labels.Italic.HasValue ||
                !labels.Italic.Value ||
                labels.FontFamily != "Arial" ||
                labels.NumberFormat != "0.0" ||
                labels.Position != "inEnd" ||
                labels.Separator != " · ")
            {
                throw new InvalidOperationException(
                    "Chart data-label options were not parsed correctly.");
            }

            XmlDocument leaderLineDoc =
                new XmlDocument();

            leaderLineDoc.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:chart><c:plotArea><c:pieChart>" +
                "<c:dLbls><c:showVal val=\"1\"/><c:dLblPos val=\"outEnd\"/><c:showLeaderLines val=\"1\"/>" +
                "<c:leaderLines><c:spPr><a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"228866\"/></a:solidFill><a:prstDash val=\"dashDot\"/></a:ln></c:spPr></c:leaderLines>" +
                "</c:dLbls></c:pieChart></c:plotArea></c:chart></c:chartSpace>");

            ChartLabelOptions leaderLabels =
                ReadChartLabelOptions(
                    leaderLineDoc);
            ChartLineStyle leaderStyle =
                ReadChartLeaderLineStyle(
                    leaderLineDoc,
                    theme);

            if (!leaderLabels.ShowLeaderLines ||
                leaderLabels.Position != "outEnd" ||
                leaderStyle.Color.R != 0x22 ||
                leaderStyle.Color.G != 0x88 ||
                leaderStyle.Color.B != 0x66 ||
                leaderStyle.Width <= 1f ||
                leaderStyle.DashStyle !=
                    DashStyle.DashDot)
            {
                throw new InvalidOperationException(
                    "Chart leader-line visibility or style was not parsed correctly.");
            }

            ChartSeriesData leaderSeries =
                new ChartSeriesData();
            leaderSeries.Name = "Only";
            leaderSeries.Categories.Add(
                "Only");
            leaderSeries.Values.Add(
                100.0);

            using (Bitmap leaderBitmap =
                new Bitmap(
                    220,
                    220,
                    PixelFormat.Format32bppArgb))
            using (Graphics leaderGraphics =
                Graphics.FromImage(
                    leaderBitmap))
            {
                leaderGraphics.Clear(
                    Color.White);

                DrawPieChartValueLabels(
                    leaderGraphics,
                    new RectangleF(
                        10f,
                        10f,
                        200f,
                        200f),
                    leaderSeries,
                    leaderLabels,
                    0f,
                    leaderStyle);

                bool leaderVisible =
                    false;

                for (int y = 106;
                     y <= 114 &&
                     !leaderVisible;
                     y++)
                {
                    for (int x = 17;
                         x <= 33;
                         x++)
                    {
                        Color pixel =
                            leaderBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.G > pixel.R + 25 &&
                            pixel.G > pixel.B + 5 &&
                            pixel.G > 70)
                        {
                            leaderVisible = true;
                            break;
                        }
                    }
                }

                if (!leaderVisible)
                {
                    throw new InvalidOperationException(
                        "Chart showLeaderLines did not render an outside-end leader line.");
                }
            }

            XmlDocument pointOverrideDoc =
                new XmlDocument();

            pointOverrideDoc.LoadXml(
                "<c:ser xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:dLbls>" +
                "<c:dLbl><c:idx val=\"0\"/><c:showVal val=\"0\"/><c:showSerName val=\"1\"/><c:dLblPos val=\"ctr\"/><c:spPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:srgbClr val=\"CCEEFF\"/></a:solidFill><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"19050\"><a:solidFill><a:srgbClr val=\"336699\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></c:spPr><c:txPr><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:pPr><a:defRPr sz=\"1100\" b=\"0\" i=\"1\"><a:latin typeface=\"Arial\"/><a:solidFill><a:srgbClr val=\"AA3377\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:separator> / </c:separator></c:dLbl>" +
                "<c:dLbl><c:idx val=\"1\"/><c:delete val=\"1\"/></c:dLbl>" +
                "</c:dLbls></c:ser>");

            XmlDocument labelScopeDoc =
                new XmlDocument();

            labelScopeDoc.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart><c:plotArea><c:barChart>" +
                "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
                "<c:dLbls><c:dLbl><c:idx val=\"0\"/><c:showSerName val=\"1\"/></c:dLbl></c:dLbls>" +
                "</c:ser>" +
                "<c:dLbls><c:showVal val=\"1\"/><c:dLblPos val=\"outEnd\"/></c:dLbls>" +
                "</c:barChart></c:plotArea></c:chart></c:chartSpace>");

            ChartLabelOptions scopedDefaults =
                ReadChartLabelOptions(
                    labelScopeDoc);

            if (!scopedDefaults.ShowValue ||
                scopedDefaults.ShowSeriesName ||
                scopedDefaults.Position != "outEnd")
            {
                throw new InvalidOperationException(
                    "Chart-level dLbls was confused with series-level dLbls.");
            }

            ChartSeriesData pointOverrideSeries =
                new ChartSeriesData();
            pointOverrideSeries.Name = "North";
            pointOverrideSeries.Values.Add(42.0);
            pointOverrideSeries.Values.Add(19.0);
            pointOverrideSeries.Categories.Add("Q1");
            pointOverrideSeries.Categories.Add("Q2");

            ReadChartSeriesLabelOverrides(
                pointOverrideDoc.DocumentElement,
                pointOverrideSeries,
                theme);

            ChartLabelOptions resolvedPointZero =
                ResolveChartPointLabelOptions(
                    labels,
                    pointOverrideSeries,
                    0);
            ChartLabelOptions resolvedPointOne =
                ResolveChartPointLabelOptions(
                    labels,
                    pointOverrideSeries,
                    1);

            if (resolvedPointZero.ShowValue ||
                !resolvedPointZero.ShowSeriesName ||
                !resolvedPointZero.TextColor.HasValue ||
                resolvedPointZero.TextColor.Value.R != 0xAA ||
                resolvedPointZero.TextColor.Value.G != 0x33 ||
                resolvedPointZero.TextColor.Value.B != 0x77 ||
                !resolvedPointZero.FillColor.HasValue ||
                resolvedPointZero.FillColor.Value.R != 0xCC ||
                resolvedPointZero.FillColor.Value.G != 0xEE ||
                resolvedPointZero.FillColor.Value.B != 0xFF ||
                resolvedPointZero.BorderStyle == null ||
                resolvedPointZero.BorderStyle.Color.R != 0x33 ||
                resolvedPointZero.BorderStyle.Color.G != 0x66 ||
                resolvedPointZero.BorderStyle.Color.B != 0x99 ||
                resolvedPointZero.BorderStyle.DashStyle !=
                    DashStyle.Dash ||
                !resolvedPointZero.FontSize.HasValue ||
                Math.Abs(
                    resolvedPointZero.FontSize.Value -
                    11f) > 0.01f ||
                !resolvedPointZero.Bold.HasValue ||
                resolvedPointZero.Bold.Value ||
                !resolvedPointZero.Italic.HasValue ||
                !resolvedPointZero.Italic.Value ||
                resolvedPointZero.FontFamily != "Arial" ||
                resolvedPointZero.Position != "ctr" ||
                resolvedPointZero.Separator != " / " ||
                resolvedPointOne.HasAny)
            {
                throw new InvalidOperationException(
                    "Per-point chart dLbl overrides were not applied correctly.");
            }

            ChartSeriesData labelSeries =
                new ChartSeriesData();
            labelSeries.Categories.Add("Q1");
            labelSeries.Values.Add(42.0);

            string formattedLabel =
                BuildChartDataLabel(
                    labels,
                    labelSeries,
                    0,
                    false);

            if (formattedLabel != "Q1 · 42.0")
            {
                throw new InvalidOperationException(
                    "Chart data-label number format was not applied: " +
                    formattedLabel);
            }

            using (Bitmap dataLabelBitmap =
                new Bitmap(
                    220,
                    120,
                    PixelFormat.Format32bppArgb))
            using (Graphics dataLabelGraphics =
                Graphics.FromImage(
                    dataLabelBitmap))
            using (Font dataLabelFont =
                SafeFont(
                    "Arial",
                    10f))
            using (Brush dataLabelBrush =
                new SolidBrush(
                    Color.Black))
            {
                dataLabelGraphics.Clear(
                    Color.White);

                DrawColumnChartDataLabel(
                    dataLabelGraphics,
                    dataLabelFont,
                    dataLabelBrush,
                    "X",
                    new RectangleF(
                        90f,
                        30f,
                        40f,
                        60f),
                    90f,
                    30f,
                    42.0,
                    "inEnd");

                bool foundInside =
                    false;

                for (int y = 31;
                     y < 52 &&
                     !foundInside;
                     y++)
                {
                    for (int x = 90;
                         x < 130;
                         x++)
                    {
                        Color pixel =
                            dataLabelBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R < 220 ||
                            pixel.G < 220 ||
                            pixel.B < 220)
                        {
                            foundInside = true;
                            break;
                        }
                    }
                }

                if (!foundInside)
                {
                    throw new InvalidOperationException(
                        "Chart dLblPos=inEnd did not move the column label inside the bar.");
                }

                dataLabelGraphics.Clear(
                    Color.White);

                ChartLabelOptions coloredLabelOptions =
                    new ChartLabelOptions();
                coloredLabelOptions.Position =
                    "ctr";
                coloredLabelOptions.TextColor =
                    Color.FromArgb(
                        170,
                        51,
                        119);
                coloredLabelOptions.FontSize =
                    14f;
                coloredLabelOptions.Bold =
                    true;
                coloredLabelOptions.Italic =
                    true;
                coloredLabelOptions.FontFamily =
                    "Arial";
                coloredLabelOptions.FillColor =
                    Color.FromArgb(
                        204,
                        238,
                        255);
                coloredLabelOptions.BorderStyle =
                    new ChartLineStyle();
                coloredLabelOptions.BorderStyle.Color =
                    Color.FromArgb(
                        51,
                        102,
                        153);
                coloredLabelOptions.BorderStyle.Width =
                    1f;
                coloredLabelOptions.BorderStyle.DashStyle =
                    DashStyle.Solid;

                DrawColumnChartDataLabel(
                    dataLabelGraphics,
                    dataLabelFont,
                    dataLabelBrush,
                    "X",
                    new RectangleF(
                        90f,
                        30f,
                        40f,
                        60f),
                    90f,
                    30f,
                    42.0,
                    coloredLabelOptions);

                bool coloredLabelVisible =
                    false;

                for (int y = 45;
                     y < 78 &&
                     !coloredLabelVisible;
                     y++)
                {
                    for (int x = 92;
                         x < 128;
                         x++)
                    {
                        Color pixel =
                            dataLabelBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R > 90 &&
                            pixel.B > 55 &&
                            pixel.R > pixel.G + 35)
                        {
                            coloredLabelVisible = true;
                            break;
                        }
                    }
                }

                if (!coloredLabelVisible)
                {
                    throw new InvalidOperationException(
                        "Chart dLbl txPr text color was not rendered.");
                }

                bool labelBoxVisible =
                    false;

                for (int y = 48;
                     y < 78 &&
                     !labelBoxVisible;
                     y++)
                {
                    for (int x = 100;
                         x < 122;
                         x++)
                    {
                        Color pixel =
                            dataLabelBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.B > 180 &&
                            pixel.G > 150 &&
                            pixel.R > 120)
                        {
                            labelBoxVisible = true;
                            break;
                        }
                    }
                }

                if (!labelBoxVisible)
                {
                    throw new InvalidOperationException(
                        "Chart dLbl spPr fill was not rendered behind the label.");
                }
            }

            if (ReadChartLegendPosition(chart) != "b")
            {
                throw new InvalidOperationException(
                    "Chart legend position was not parsed correctly.");
            }

            HashSet<int> hiddenLegendEntries =
                ReadChartLegendHiddenEntries(
                    chart);

            if (hiddenLegendEntries.Count != 1 ||
                !hiddenLegendEntries.Contains(
                    1))
            {
                throw new InvalidOperationException(
                    "Chart legendEntry delete was not parsed correctly.");
            }

            ChartLabelOptions legendTextStyle =
                ReadChartLegendTextStyle(
                    chart,
                    theme);

            if (!legendTextStyle.TextColor.HasValue ||
                legendTextStyle.TextColor.Value.R != 0xCC ||
                legendTextStyle.TextColor.Value.G != 0x22 ||
                legendTextStyle.TextColor.Value.B != 0x77 ||
                !legendTextStyle.FontSize.HasValue ||
                Math.Abs(
                    legendTextStyle.FontSize.Value -
                    12f) > 0.01f ||
                !legendTextStyle.Bold.HasValue ||
                !legendTextStyle.Bold.Value ||
                !legendTextStyle.Italic.HasValue ||
                !legendTextStyle.Italic.Value ||
                legendTextStyle.FontFamily != "Arial")
            {
                throw new InvalidOperationException(
                    "Chart legend txPr text style was not parsed correctly.");
            }

            using (Bitmap legendBitmap =
                new Bitmap(
                    320,
                    180,
                    PixelFormat.Format32bppArgb))
            using (Graphics legendGraphics =
                Graphics.FromImage(
                    legendBitmap))
            {
                legendGraphics.Clear(
                    Color.White);

                List<ChartSeriesData> legendSeries =
                    new List<ChartSeriesData>();

                ChartSeriesData legendItem =
                    new ChartSeriesData();
                legendItem.Name =
                    "Styled Legend";
                legendItem.Values.Add(
                    1.0);
                legendSeries.Add(
                    legendItem);

                ChartSeriesData hiddenLegendItem =
                    new ChartSeriesData();
                hiddenLegendItem.Name =
                    "Hidden Legend";
                hiddenLegendItem.Values.Add(
                    2.0);
                legendSeries.Add(
                    hiddenLegendItem);

                DrawChartLegend(
                    legendGraphics,
                    new RectangleF(
                        0f,
                        0f,
                        320f,
                        180f),
                    legendSeries,
                    new Color[]
                    {
                        Color.FromArgb(
                            70,
                            120,
                            205),
                        Color.FromArgb(
                            220,
                            55,
                            45)
                    },
                    "column",
                    "b",
                    legendTextStyle,
                    hiddenLegendEntries);

                bool legendStyleVisible =
                    false;

                for (int y = 140;
                     y < 178 &&
                     !legendStyleVisible;
                     y++)
                {
                    for (int x = 24;
                         x < 190;
                         x++)
                    {
                        Color pixel =
                            legendBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R >
                                pixel.G + 85 &&
                            pixel.B >
                                pixel.G + 35 &&
                            pixel.R > 120)
                        {
                            legendStyleVisible =
                                true;
                            break;
                        }
                    }
                }

                if (!legendStyleVisible)
                {
                    throw new InvalidOperationException(
                        "Chart legend txPr text color was not rendered.");
                }

                bool hiddenRedSwatchVisible =
                    false;

                for (int y = 138;
                     y < 178 &&
                     !hiddenRedSwatchVisible;
                     y++)
                {
                    for (int x = 8;
                         x < 300;
                         x++)
                    {
                        Color pixel =
                            legendBitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R > 175 &&
                            pixel.G < 105 &&
                            pixel.B < 105)
                        {
                            hiddenRedSwatchVisible =
                                true;
                            break;
                        }
                    }
                }

                if (hiddenRedSwatchVisible)
                {
                    throw new InvalidOperationException(
                        "Chart legendEntry delete still rendered the hidden legend swatch.");
                }
            }

            XmlDocument stackedChart =
                new XmlDocument();

            stackedChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart><c:plotArea><c:barChart>" +
                "<c:barDir val=\"col\"/><c:grouping val=\"percentStacked\"/>" +
                "<c:gapWidth val=\"80\"/><c:overlap val=\"100\"/>" +
                "</c:barChart></c:plotArea></c:chart></c:chartSpace>");

            ChartBarOptions stackedOptions =
                ReadChartBarOptions(
                    stackedChart);

            if (!stackedOptions.IsStacked ||
                !stackedOptions.IsPercentStacked ||
                Math.Abs(
                    stackedOptions.GapWidth -
                    80f) > 0.01f ||
                Math.Abs(
                    stackedOptions.Overlap -
                    100f) > 0.01f)
            {
                throw new InvalidOperationException(
                    "Stacked chart grouping/gap/overlap options were not parsed correctly.");
            }

            XmlDocument doughnutOptions =
                new XmlDocument();

            doughnutOptions.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart><c:plotArea><c:doughnutChart>" +
                "<c:firstSliceAng val=\"120\"/><c:holeSize val=\"72\"/>" +
                "</c:doughnutChart></c:plotArea></c:chart></c:chartSpace>");

            if (Math.Abs(
                    ReadChartFirstSliceAngle(
                        doughnutOptions) -
                    30f) > 0.01f ||
                Math.Abs(
                    ReadDoughnutHoleRatio(
                        doughnutOptions) -
                    0.72f) > 0.001f)
            {
                throw new InvalidOperationException(
                    "Pie/doughnut rotation or hole-size options were not parsed correctly.");
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

                bool titleStyleVisible =
                    false;

                for (int y = 24;
                     y < 82 &&
                     !titleStyleVisible;
                     y++)
                {
                    for (int x = 300;
                         x < 700;
                         x++)
                    {
                        Color pixel =
                            bitmap.GetPixel(
                                x,
                                y);

                        if (pixel.G >
                                pixel.R + 55 &&
                            pixel.G >
                                pixel.B + 35 &&
                            pixel.G > 90)
                        {
                            titleStyleVisible =
                                true;
                            break;
                        }
                    }
                }

                if (!titleStyleVisible)
                {
                    throw new InvalidOperationException(
                        "Chart title txPr text color was not rendered.");
                }

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

        private static void ValidateSyntheticTextInheritance()
        {
            Dictionary<string, PlaceholderTextStyleContext>
                previousPlaceholders =
                    activePlaceholderTextStyles;

            PlaceholderTextStyleContext previousDefault =
                activePresentationDefaultTextStyle;
            PlaceholderTextStyleContext previousTitle =
                activeMasterTitleStyle;
            PlaceholderTextStyleContext previousBody =
                activeMasterBodyStyle;
            PlaceholderTextStyleContext previousOther =
                activeMasterOtherStyle;

            try
            {
                activePlaceholderTextStyles =
                    new Dictionary<string, PlaceholderTextStyleContext>(
                        StringComparer.OrdinalIgnoreCase);
                activePresentationDefaultTextStyle =
                    new PlaceholderTextStyleContext();
                activeMasterTitleStyle =
                    new PlaceholderTextStyleContext();
                activeMasterBodyStyle =
                    new PlaceholderTextStyleContext();
                activeMasterOtherStyle =
                    new PlaceholderTextStyleContext();

                XmlDocument source =
                    new XmlDocument();

                source.LoadXml(
                    "<root xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<a:lvl2pPr marL=\"300000\"><a:defRPr sz=\"1700\"/></a:lvl2pPr>" +
                    "<a:master marR=\"200000\"><a:defRPr b=\"1\"/></a:master>" +
                    "<a:layout spcAft=\"0\"><a:defRPr i=\"1\"/></a:layout>" +
                    "<a:bodyPr lIns=\"120000\" tIns=\"60000\"><a:normAutofit fontScale=\"90000\"/></a:bodyPr>" +
                    "</root>");

                activePresentationDefaultTextStyle
                    .Levels[1]
                    .Add(
                        source.DocumentElement
                            .ChildNodes[0]
                            .CloneNode(true));

                activeMasterBodyStyle
                    .Levels[1]
                    .Add(
                        source.DocumentElement
                            .ChildNodes[1]
                            .CloneNode(true));

                PlaceholderTextStyleContext specific =
                    new PlaceholderTextStyleContext();

                specific.Levels[1].Add(
                    source.DocumentElement
                        .ChildNodes[2]
                        .CloneNode(true));

                specific.BodyProperties.Add(
                    source.DocumentElement
                        .ChildNodes[3]
                        .CloneNode(true));

                activePlaceholderTextStyles.Add(
                    "type:body",
                    specific);

                XmlDocument shapeDocument =
                    new XmlDocument();

                shapeDocument.LoadXml(
                    "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<p:nvSpPr><p:nvPr><p:ph type=\"body\"/></p:nvPr></p:nvSpPr>" +
                    "<p:txBody>" +
                    "<a:bodyPr anchor=\"ctr\"/>" +
                    "<a:lstStyle><a:lvl2pPr indent=\"-100000\"/></a:lstStyle>" +
                    "<a:p><a:pPr lvl=\"1\" algn=\"r\"/><a:r><a:t>Inherited</a:t></a:r></a:p>" +
                    "</p:txBody></p:sp>");

                XmlNode resolved =
                    BuildRichInheritedShape(
                        shapeDocument.DocumentElement);

                XmlNode bodyPr =
                    FindFirst(
                        resolved,
                        "bodyPr");

                if (bodyPr == null ||
                    GetAttr(bodyPr, "lIns") != "120000" ||
                    GetAttr(bodyPr, "tIns") != "60000" ||
                    GetAttr(bodyPr, "anchor") != "ctr" ||
                    DirectChild(bodyPr, "normAutofit") == null)
                {
                    throw new InvalidOperationException(
                        "Inherited placeholder bodyPr properties were not merged.");
                }

                XmlNode paragraph =
                    FindFirst(
                        resolved,
                        "p");
                XmlNode pPr =
                    DirectChild(
                        paragraph,
                        "pPr");
                XmlNode defRPr =
                    pPr != null
                        ? DirectChild(
                            pPr,
                            "defRPr")
                        : null;

                if (pPr == null ||
                    GetAttr(pPr, "marL") != "300000" ||
                    GetAttr(pPr, "marR") != "200000" ||
                    GetAttr(pPr, "indent") != "-100000" ||
                    GetAttr(pPr, "algn") != "r" ||
                    defRPr == null ||
                    GetAttr(defRPr, "sz") != "1700" ||
                    GetAttr(defRPr, "b") != "1" ||
                    GetAttr(defRPr, "i") != "1")
                {
                    throw new InvalidOperationException(
                        "Presentation/master/layout/list/local text inheritance order was not preserved.");
                }
            }
            finally
            {
                activePlaceholderTextStyles =
                    previousPlaceholders;
                activePresentationDefaultTextStyle =
                    previousDefault;
                activeMasterTitleStyle =
                    previousTitle;
                activeMasterBodyStyle =
                    previousBody;
                activeMasterOtherStyle =
                    previousOther;
            }
        }

        private static void ValidateSyntheticColorTransforms()
        {
            XmlDocument document =
                new XmlDocument();

            document.LoadXml(
                "<a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<a:srgbClr val=\"CC6633\">" +
                "<a:hueOff val=\"7200000\"/>" +
                "<a:satMod val=\"55000\"/>" +
                "<a:lumOff val=\"8000\"/>" +
                "</a:srgbClr>" +
                "</a:solidFill>");

            Color? transformed =
                ReadColorFromFill(
                    document.DocumentElement,
                    new Dictionary<string, Color>(
                        StringComparer.OrdinalIgnoreCase));

            if (!transformed.HasValue)
            {
                throw new InvalidOperationException(
                    "DrawingML HSL color transform did not produce a color.");
            }

            Color original =
                Color.FromArgb(
                    0xCC,
                    0x66,
                    0x33);

            if (Math.Abs(
                    transformed.Value.R -
                    original.R) < 8 &&
                Math.Abs(
                    transformed.Value.G -
                    original.G) < 8 &&
                Math.Abs(
                    transformed.Value.B -
                    original.B) < 8)
            {
                throw new InvalidOperationException(
                    "DrawingML hue/saturation/luminance transforms were not applied.");
            }

            XmlDocument grayDocument =
                new XmlDocument();

            grayDocument.LoadXml(
                "<a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<a:srgbClr val=\"2080E0\"><a:gray/></a:srgbClr>" +
                "</a:solidFill>");

            Color? gray =
                ReadColorFromFill(
                    grayDocument.DocumentElement,
                    null);

            if (!gray.HasValue ||
                Math.Abs(
                    gray.Value.R -
                    gray.Value.G) > 1 ||
                Math.Abs(
                    gray.Value.G -
                    gray.Value.B) > 1)
            {
                throw new InvalidOperationException(
                    "DrawingML gray transform was not applied.");
            }

            XmlDocument inverseDocument =
                new XmlDocument();

            inverseDocument.LoadXml(
                "<a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<a:srgbClr val=\"102030\"><a:inv/></a:srgbClr>" +
                "</a:solidFill>");

            Color? inverse =
                ReadColorFromFill(
                    inverseDocument.DocumentElement,
                    null);

            if (!inverse.HasValue ||
                inverse.Value.R != 0xEF ||
                inverse.Value.G != 0xDF ||
                inverse.Value.B != 0xCF)
            {
                throw new InvalidOperationException(
                    "DrawingML inverse color transform was not applied.");
            }
        }

        private static void ValidateSyntheticShapeEffects(
            string outputPath)
        {
            XmlDocument shapeProperties =
                new XmlDocument();

            shapeProperties.LoadXml(
                "<a:spPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<a:solidFill><a:srgbClr val=\"3978D4\"/></a:solidFill>" +
                "<a:effectLst>" +
                "<a:outerShdw dist=\"80000\" dir=\"2700000\"><a:srgbClr val=\"202020\"><a:alpha val=\"45000\"/></a:srgbClr></a:outerShdw>" +
                "<a:glow rad=\"110000\"><a:srgbClr val=\"55A8FF\"><a:alpha val=\"65000\"/></a:srgbClr></a:glow>" +
                "<a:softEdge rad=\"60000\"/>" +
                "<a:reflection dist=\"50000\" sy=\"70000\" stA=\"42000\" endA=\"0\"/>" +
                "<a:innerShdw blurRad=\"50000\" dist=\"24000\" dir=\"13500000\"><a:srgbClr val=\"101010\"><a:alpha val=\"50000\"/></a:srgbClr></a:innerShdw>" +
                "</a:effectLst>" +
                "<a:sp3d><a:bevelT w=\"90000\" h=\"70000\"/><a:bevelB w=\"50000\" h=\"50000\"/></a:sp3d>" +
                "</a:spPr>");

            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            RectangleF rect =
                new RectangleF(
                    130f,
                    80f,
                    220f,
                    110f);

            using (Bitmap bitmap =
                new Bitmap(
                    520,
                    380,
                    PixelFormat.Format32bppArgb))
            using (Graphics graphics =
                Graphics.FromImage(bitmap))
            using (GraphicsPath path =
                new GraphicsPath())
            {
                graphics.Clear(
                    Color.White);
                graphics.SmoothingMode =
                    SmoothingMode.AntiAlias;

                path.AddRectangle(
                    rect);

                DrawShapeVisualEffects(
                    graphics,
                    shapeProperties.DocumentElement,
                    path,
                    rect,
                    theme);

                using (Brush fill =
                    new SolidBrush(
                        Color.FromArgb(
                            57,
                            120,
                            212)))
                {
                    graphics.FillPath(
                        fill,
                        path);
                }

                DrawShapePostFillEffects(
                    graphics,
                    shapeProperties.DocumentElement,
                    path,
                    rect,
                    theme);

                Color glowPixel =
                    bitmap.GetPixel(
                        120,
                        120);

                Color reflectionPixel =
                    bitmap.GetPixel(
                        240,
                        215);

                if (glowPixel.R > 247 &&
                    glowPixel.G > 247 &&
                    glowPixel.B > 247)
                {
                    throw new InvalidOperationException(
                        "Shape glow/soft-edge approximation did not render outside the source path.");
                }

                if (reflectionPixel.R > 247 &&
                    reflectionPixel.G > 247 &&
                    reflectionPixel.B > 247)
                {
                    throw new InvalidOperationException(
                        "Shape reflection approximation did not render below the source path.");
                }

                HashSet<int> interiorBuckets =
                    new HashSet<int>();

                for (int y = (int)rect.Top + 3;
                     y < (int)rect.Bottom - 3;
                     y += 6)
                {
                    for (int x = (int)rect.Left + 3;
                         x < (int)rect.Right - 3;
                         x += 6)
                    {
                        Color pixel =
                            bitmap.GetPixel(
                                x,
                                y);

                        interiorBuckets.Add(
                            ((pixel.R / 24) << 12) |
                            ((pixel.G / 24) << 6) |
                            (pixel.B / 24));
                    }
                }

                if (interiorBuckets.Count < 3)
                {
                    throw new InvalidOperationException(
                        "Inner-shadow / bevel post-fill approximation did not alter the shape interior.");
                }

                XmlDocument presetShadow =
                    new XmlDocument();

                presetShadow.LoadXml(
                    "<a:spPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<a:solidFill><a:srgbClr val=\"55AA77\"/></a:solidFill>" +
                    "<a:effectLst><a:prstShdw prst=\"shdw1\" dist=\"70000\" dir=\"2700000\">" +
                    "<a:srgbClr val=\"202020\"><a:alpha val=\"50000\"/></a:srgbClr>" +
                    "</a:prstShdw></a:effectLst></a:spPr>");

                RectangleF presetRect =
                    new RectangleF(
                        385f,
                        80f,
                        95f,
                        72f);

                using (GraphicsPath presetPath =
                    new GraphicsPath())
                {
                    presetPath.AddRectangle(
                        presetRect);

                    DrawShapeVisualEffects(
                        graphics,
                        presetShadow.DocumentElement,
                        presetPath,
                        presetRect,
                        theme);

                    using (Brush presetFill =
                        new SolidBrush(
                            Color.FromArgb(
                                85,
                                170,
                                119)))
                    {
                        graphics.FillPath(
                            presetFill,
                            presetPath);
                    }
                }

                Color presetShadowPixel =
                    bitmap.GetPixel(
                        486,
                        158);

                if (presetShadowPixel.R > 247 &&
                    presetShadowPixel.G > 247 &&
                    presetShadowPixel.B > 247)
                {
                    throw new InvalidOperationException(
                        "Preset shadow approximation did not render outside the source shape.");
                }

                XmlDocument extrusionOnly =
                    new XmlDocument();

                extrusionOnly.LoadXml(
                    "<a:spPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<a:solidFill><a:srgbClr val=\"4D78B8\"/></a:solidFill>" +
                    "<a:scene3d><a:camera prst=\"perspectiveFront\"><a:rot lat=\"900000\" lon=\"-1200000\" rev=\"0\"/></a:camera>" +
                    "<a:lightRig rig=\"balanced\" dir=\"tr\"/></a:scene3d>" +
                    "<a:sp3d extrusionH=\"110000\"><a:extrusionClr><a:srgbClr val=\"203858\"/></a:extrusionClr></a:sp3d>" +
                    "</a:spPr>");

                RectangleF extrusionRect =
                    new RectangleF(
                        385f,
                        255f,
                        82f,
                        52f);

                using (GraphicsPath extrusionPath =
                    new GraphicsPath())
                {
                    extrusionPath.AddRectangle(
                        extrusionRect);

                    DrawShapeVisualEffects(
                        graphics,
                        extrusionOnly.DocumentElement,
                        extrusionPath,
                        extrusionRect,
                        theme);

                    using (Brush extrusionFill =
                        new SolidBrush(
                            Color.FromArgb(
                                77,
                                120,
                                184)))
                    {
                        graphics.FillPath(
                            extrusionFill,
                            extrusionPath);
                    }
                }

                int extrusionPixels = 0;

                for (int y = 245;
                     y < 330;
                     y += 3)
                {
                    for (int x = 375;
                         x < 505;
                         x += 3)
                    {
                        bool insideFront =
                            x >= extrusionRect.Left &&
                            x <= extrusionRect.Right &&
                            y >= extrusionRect.Top &&
                            y <= extrusionRect.Bottom;

                        if (insideFront)
                            continue;

                        Color pixel =
                            bitmap.GetPixel(
                                x,
                                y);

                        if (pixel.R < 235 ||
                            pixel.G < 235 ||
                            pixel.B < 235)
                        {
                            extrusionPixels++;
                        }
                    }
                }

                if (extrusionPixels < 4)
                {
                    throw new InvalidOperationException(
                        "3D extrusion approximation did not render depth outside the front face.");
                }

                bitmap.Save(
                    outputPath,
                    ImageFormat.Png);
            }

            RequireDiagnosticFile(
                outputPath);
        }

        private static void ValidateSyntheticImageEffects(
            string outputPath)
        {
            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            using (Bitmap source =
                new Bitmap(
                    3,
                    1,
                    PixelFormat.Format32bppArgb))
            {
                source.SetPixel(
                    0,
                    0,
                    Color.Black);
                source.SetPixel(
                    1,
                    0,
                    Color.FromArgb(
                        128,
                        128,
                        128));
                source.SetPixel(
                    2,
                    0,
                    Color.White);

                XmlDocument duotone =
                    new XmlDocument();

                duotone.LoadXml(
                    "<a:blip xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<a:duotone>" +
                    "<a:srgbClr val=\"102040\"/>" +
                    "<a:srgbClr val=\"F0C040\"/>" +
                    "</a:duotone>" +
                    "</a:blip>");

                using (Bitmap output =
                    new Bitmap(
                        640,
                        220,
                        PixelFormat.Format32bppArgb))
                using (Graphics graphics =
                    Graphics.FromImage(output))
                {
                    graphics.Clear(
                        Color.White);
                    graphics.InterpolationMode =
                        InterpolationMode.NearestNeighbor;
                    graphics.PixelOffsetMode =
                        PixelOffsetMode.Half;

                    DrawImageWithDrawingEffects(
                        graphics,
                        source,
                        new RectangleF(
                            0f,
                            0f,
                            3f,
                            1f),
                        new RectangleF(
                            20f,
                            20f,
                            300f,
                            80f),
                        duotone.DocumentElement,
                        theme);

                    Color dark =
                        output.GetPixel(
                            55,
                            55);
                    Color light =
                        output.GetPixel(
                            285,
                            55);

                    if (dark.B < 35 ||
                        light.R < 190 ||
                        Math.Abs(
                            dark.R -
                            light.R) < 70)
                    {
                        throw new InvalidOperationException(
                            "Duotone image effect did not map luminance to the configured colors.");
                    }

                    using (Bitmap replacementSource =
                        new Bitmap(
                            1,
                            1,
                            PixelFormat.Format32bppArgb))
                    {
                        replacementSource.SetPixel(
                            0,
                            0,
                            Color.FromArgb(
                                255,
                                0,
                                0));

                        XmlDocument change =
                            new XmlDocument();

                        change.LoadXml(
                            "<a:blip xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                            "<a:clrChange>" +
                            "<a:clrFrom><a:srgbClr val=\"FF0000\"/></a:clrFrom>" +
                            "<a:clrTo><a:srgbClr val=\"22AA66\"/></a:clrTo>" +
                            "</a:clrChange>" +
                            "</a:blip>");

                        DrawImageWithDrawingEffects(
                            graphics,
                            replacementSource,
                            new RectangleF(
                                0f,
                                0f,
                                1f,
                                1f),
                            new RectangleF(
                                350f,
                                20f,
                                220f,
                                80f),
                            change.DocumentElement,
                            theme);
                    }

                    Color changed =
                        output.GetPixel(
                            460,
                            55);

                    if (changed.G < 130 ||
                        changed.R > 90 ||
                        changed.B < 60)
                    {
                        throw new InvalidOperationException(
                            "clrChange image effect did not remap the configured source color.");
                    }

                    using (Bitmap edgeSource =
                        new Bitmap(
                            24,
                            8,
                            PixelFormat.Format32bppArgb))
                    {
                        for (int y = 0;
                             y < edgeSource.Height;
                             y++)
                        {
                            for (int x = 0;
                                 x < edgeSource.Width;
                                 x++)
                            {
                                edgeSource.SetPixel(
                                    x,
                                    y,
                                    x < 12
                                        ? Color.Black
                                        : Color.White);
                            }
                        }

                        XmlDocument blurDocument =
                            new XmlDocument();

                        blurDocument.LoadXml(
                            "<a:blur xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" rad=\"127000\" grow=\"0\"/>");

                        using (Bitmap blurred =
                            CreateBlurredImageApproximation(
                                edgeSource,
                                blurDocument.DocumentElement))
                        {
                            if (blurred == null)
                            {
                                throw new InvalidOperationException(
                                    "DrawingML blur approximation did not create an image.");
                            }

                            bool foundIntermediate =
                                false;

                            for (int x = 4;
                                 x < 20;
                                 x++)
                            {
                                Color sample =
                                    blurred.GetPixel(
                                        x,
                                        4);

                                if (sample.R > 16 &&
                                    sample.R < 239)
                                {
                                    foundIntermediate =
                                        true;
                                    break;
                                }
                            }

                            Color farLeft =
                                blurred.GetPixel(
                                    0,
                                    4);
                            Color farRight =
                                blurred.GetPixel(
                                    23,
                                    4);

                            if (!foundIntermediate ||
                                farLeft.R > 120 ||
                                farRight.R < 135)
                            {
                                throw new InvalidOperationException(
                                    "DrawingML blur approximation did not soften a hard image edge.");
                            }
                        }

                        XmlDocument artisticBlurDocument =
                            new XmlDocument();

                        artisticBlurDocument.LoadXml(
                            "<a14:artisticBlur xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" radius=\"40\"/>");

                        using (Bitmap artisticBlurred =
                            CreateArtisticBlurImageApproximation(
                                edgeSource,
                                artisticBlurDocument.DocumentElement))
                        {
                            if (artisticBlurred == null)
                            {
                                throw new InvalidOperationException(
                                    "Office 2010 artistic blur approximation did not create an image.");
                            }

                            bool artisticIntermediate =
                                false;

                            for (int x = 4;
                                 x < 20;
                                 x++)
                            {
                                int red =
                                    artisticBlurred.GetPixel(
                                        x,
                                        4).R;

                                if (red > 16 &&
                                    red < 239)
                                {
                                    artisticIntermediate =
                                        true;
                                    break;
                                }
                            }

                            if (!artisticIntermediate)
                            {
                                throw new InvalidOperationException(
                                    "Office 2010 artistic blur approximation did not soften the source edge.");
                            }
                        }
                    }

                    using (Bitmap impulse =
                        new Bitmap(
                            9,
                            9,
                            PixelFormat.Format32bppArgb))
                    {
                        using (Graphics impulseGraphics =
                            Graphics.FromImage(
                                impulse))
                        {
                            impulseGraphics.Clear(
                                Color.FromArgb(
                                    100,
                                    100,
                                    100));
                        }

                        impulse.SetPixel(
                            4,
                            4,
                            Color.FromArgb(
                                160,
                                160,
                                160));

                        XmlDocument sharpenDocument =
                            new XmlDocument();
                        sharpenDocument.LoadXml(
                            "<a14:sharpenSoften xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" amount=\"75000\"/>");

                        XmlDocument softenDocument =
                            new XmlDocument();
                        softenDocument.LoadXml(
                            "<a14:sharpenSoften xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" amount=\"-75000\"/>");

                        using (Bitmap sharpened =
                            CreateSharpenSoftenImageApproximation(
                                impulse,
                                sharpenDocument.DocumentElement))
                        using (Bitmap softened =
                            CreateSharpenSoftenImageApproximation(
                                impulse,
                                softenDocument.DocumentElement))
                        {
                            if (sharpened == null ||
                                softened == null)
                            {
                                throw new InvalidOperationException(
                                    "Office 2010 sharpen/soften effect did not create output images.");
                            }

                            int sharpenedCenter =
                                sharpened.GetPixel(
                                    4,
                                    4).R;
                            int softenedCenter =
                                softened.GetPixel(
                                    4,
                                    4).R;

                            if (sharpenedCenter <= 160 ||
                                softenedCenter >= 160 ||
                                sharpenedCenter <=
                                    softenedCenter)
                            {
                                throw new InvalidOperationException(
                                    "Office 2010 sharpen/soften effect did not adjust local contrast as expected.");
                            }
                        }
                    }

                    output.Save(
                        outputPath,
                        ImageFormat.Png);
                }
            }

            RequireDiagnosticFile(
                outputPath);
        }

        private static void ValidateSyntheticCustomGeometry(
            string outputPath)
        {
            XmlDocument document =
                new XmlDocument();

            document.LoadXml(
                "<a:custGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<a:avLst><a:gd name=\"adj\" fmla=\"val 40000\"/></a:avLst>" +
                "<a:gdLst>" +
                "<a:gd name=\"x1\" fmla=\"*/ w adj 100000\"/>" +
                "<a:gd name=\"x2\" fmla=\"+- w 0 x1\"/>" +
                "<a:gd name=\"ym\" fmla=\"*/ h 1 2\"/>" +
                "</a:gdLst>" +
                "<a:pathLst>" +
                "<a:path w=\"21600\" h=\"21600\">" +
                "<a:moveTo><a:pt x=\"0\" y=\"ym\"/></a:moveTo>" +
                "<a:lnTo><a:pt x=\"x1\" y=\"0\"/></a:lnTo>" +
                "<a:lnTo><a:pt x=\"x2\" y=\"0\"/></a:lnTo>" +
                "<a:lnTo><a:pt x=\"w\" y=\"ym\"/></a:lnTo>" +
                "<a:lnTo><a:pt x=\"x2\" y=\"h\"/></a:lnTo>" +
                "<a:lnTo><a:pt x=\"x1\" y=\"h\"/></a:lnTo>" +
                "<a:close/>" +
                "</a:path>" +
                "<a:path w=\"21600\" h=\"21600\" fill=\"none\">" +
                "<a:moveTo><a:pt x=\"hc\" y=\"0\"/></a:moveTo>" +
                "<a:arcTo wR=\"wd2\" hR=\"hd2\" stAng=\"3cd4\" swAng=\"cd2\"/>" +
                "</a:path>" +
                "</a:pathLst>" +
                "</a:custGeom>");

            RectangleF target =
                new RectangleF(
                    20f,
                    20f,
                    560f,
                    360f);

            using (GraphicsPath path =
                BuildCustomGeometryPath(
                    document.DocumentElement,
                    target))
            {
                if (path == null ||
                    path.PointCount < 8)
                {
                    throw new InvalidOperationException(
                        "Custom geometry guide/arc path was not created.");
                }

                RectangleF bounds =
                    path.GetBounds();

                if (bounds.Width <
                        target.Width * 0.85f ||
                    bounds.Height <
                        target.Height * 0.85f)
                {
                    throw new InvalidOperationException(
                        "Custom geometry guide evaluation produced an unexpected path extent.");
                }

                using (Bitmap bitmap =
                    new Bitmap(
                        600,
                        400,
                        PixelFormat.Format32bppArgb))
                using (Graphics graphics =
                    Graphics.FromImage(bitmap))
                {
                    graphics.Clear(
                        Color.White);
                    graphics.SmoothingMode =
                        System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                    using (Brush fill =
                        new SolidBrush(
                            Color.FromArgb(
                                224,
                                236,
                                249)))
                    using (Pen line =
                        new Pen(
                            Color.FromArgb(
                                54,
                                104,
                                171),
                            3f))
                    {
                        graphics.FillPath(
                            fill,
                            path);
                        graphics.DrawPath(
                            line,
                            path);
                    }

                    bitmap.Save(
                        outputPath,
                        ImageFormat.Png);
                }
            }

            RequireDiagnosticFile(
                outputPath);
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

                int tabFarRight =
                    CountNonWhiteDiagnosticPixels(
                        bitmap,
                        new Rectangle(
                            430,
                            20,
                            240,
                            100));

                if (tabFarRight < 2)
                {
                    throw new InvalidOperationException(
                        "Right/decimal tab alignment did not place content in the expected far-right region.");
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
                "<a:tab pos=\"1371600\" algn=\"ctr\"/>" +
                "<a:tab pos=\"2743200\" algn=\"r\"/>" +
                "<a:tab pos=\"3657600\" algn=\"dec\"/>" +
                "</a:tabLst></a:pPr>" +
                "<a:r><a:rPr sz=\"1800\" b=\"1\"/><a:t>Alpha</a:t></a:r>" +
                "<a:tab/>" +
                "<a:r><a:rPr sz=\"1800\"/><a:t>Beta</a:t></a:r>" +
                "<a:tab/>" +
                "<a:r><a:rPr sz=\"1800\"/><a:t>Gamma</a:t></a:r>" +
                "<a:tab/>" +
                "<a:r><a:rPr sz=\"1800\"/><a:t>123.45</a:t></a:r>" +
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
                "<c:spPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:solidFill><a:srgbClr val=\"");
            xml.Append(color);
            xml.Append(
                "\"/></a:solidFill><a:ln w=\"38100\"><a:prstDash val=\"dash\"/></a:ln></c:spPr>");
            xml.Append("<c:marker><c:symbol val=\"diamond\"/><c:size val=\"9\"/></c:marker>");

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

        private static void ValidateSyntheticExtendedCharts(
            string outputPath)
        {
            Dictionary<string, Color> theme =
                new Dictionary<string, Color>(
                    StringComparer.OrdinalIgnoreCase);

            theme["accent1"] =
                Color.FromArgb(
                    74,
                    122,
                    206);
            theme["accent2"] =
                Color.FromArgb(
                    210,
                    86,
                    72);
            theme["accent3"] =
                Color.FromArgb(
                    92,
                    156,
                    90);
            theme["accent4"] =
                Color.FromArgb(
                    138,
                    103,
                    176);

            XmlDocument overlayChart =
                new XmlDocument();

            overlayChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Overlay</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"1\"/></c:title>" +
                "<c:plotArea/>" +
                "<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"1\"/></c:legend>" +
                "</c:chart></c:chartSpace>");

            if (!ReadChartOverlay(
                    overlayChart,
                    "title") ||
                !ReadChartOverlay(
                    overlayChart,
                    "legend"))
            {
                throw new InvalidOperationException(
                    "Chart title or legend overlay setting was not parsed.");
            }

            using (Bitmap overlayBitmap =
                new Bitmap(
                    360,
                    240,
                    PixelFormat.Format32bppArgb))
            using (Graphics overlayGraphics =
                Graphics.FromImage(
                    overlayBitmap))
            {
                overlayGraphics.Clear(
                    Color.White);

                RectangleF overlayPlot;

                PrepareChartSurface(
                    overlayGraphics,
                    new RectangleF(
                        10f,
                        10f,
                        320f,
                        200f),
                    overlayChart,
                    theme,
                    out overlayPlot);

                if (overlayPlot.Top > 30f ||
                    overlayPlot.Right < 295f)
                {
                    throw new InvalidOperationException(
                        "Chart overlay title or legend still consumed plot layout space.");
                }
            }

            XmlDocument axisTitleOnlyChart =
                new XmlDocument();

            axisTitleOnlyChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:chart><c:plotArea><c:catAx>" +
                "<c:title><c:tx><c:rich><a:p><a:r><a:t>Axis Only</a:t></a:r></a:p></c:rich></c:tx>" +
                "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val=\"CC2244\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>" +
                "</c:title>" +
                "</c:catAx></c:plotArea></c:chart></c:chartSpace>");

            ChartLabelOptions axisOnlyChartTitleStyle =
                ReadChartTitleTextStyle(
                    axisTitleOnlyChart,
                    theme);

            if (!string.IsNullOrEmpty(
                    ReadChartTitle(
                        axisTitleOnlyChart)) ||
                ReadChartOverlay(
                    axisTitleOnlyChart,
                    "title") ||
                (axisOnlyChartTitleStyle != null &&
                 axisOnlyChartTitleStyle.TextColor.HasValue))
            {
                throw new InvalidOperationException(
                    "Axis title leaked into chart-level title scope.");
            }

            XmlDocument manualLayoutChart =
                new XmlDocument();

            manualLayoutChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart><c:plotArea><c:layout><c:manualLayout>" +
                "<c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/>" +
                "<c:x val=\"0.18\"/><c:y val=\"0.22\"/>" +
                "<c:w val=\"0.58\"/><c:h val=\"0.52\"/>" +
                "</c:manualLayout></c:layout></c:plotArea></c:chart>" +
                "</c:chartSpace>");

            using (Bitmap manualLayoutBitmap =
                new Bitmap(
                    360,
                    240,
                    PixelFormat.Format32bppArgb))
            using (Graphics manualLayoutGraphics =
                Graphics.FromImage(
                    manualLayoutBitmap))
            {
                manualLayoutGraphics.Clear(
                    Color.White);

                RectangleF manualPlot;

                PrepareChartSurface(
                    manualLayoutGraphics,
                    new RectangleF(
                        10f,
                        10f,
                        320f,
                        200f),
                    manualLayoutChart,
                    theme,
                    out manualPlot);

                if (Math.Abs(
                        manualPlot.Left -
                        67.6f) > 1.0f ||
                    Math.Abs(
                        manualPlot.Top -
                        54.0f) > 1.0f ||
                    Math.Abs(
                        manualPlot.Width -
                        185.6f) > 1.0f ||
                    Math.Abs(
                        manualPlot.Height -
                        104.0f) > 1.0f)
                {
                    throw new InvalidOperationException(
                        "Chart plotArea manual layout coordinates were not applied.");
                }
            }

            XmlDocument innerTargetChart =
                new XmlDocument();

            innerTargetChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart><c:plotArea><c:layout><c:manualLayout>" +
                "<c:layoutTarget val=\"inner\"/>" +
                "<c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/>" +
                "<c:x val=\"0.10\"/><c:y val=\"0.10\"/>" +
                "<c:w val=\"0.60\"/><c:h val=\"0.60\"/>" +
                "</c:manualLayout></c:layout>" +
                "<c:catAx/><c:valAx/>" +
                "</c:plotArea></c:chart></c:chartSpace>");

            XmlDocument outerTargetChart =
                new XmlDocument();

            outerTargetChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart><c:plotArea><c:layout><c:manualLayout>" +
                "<c:layoutTarget val=\"outer\"/>" +
                "<c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/>" +
                "<c:x val=\"0.10\"/><c:y val=\"0.10\"/>" +
                "<c:w val=\"0.60\"/><c:h val=\"0.60\"/>" +
                "</c:manualLayout></c:layout>" +
                "<c:catAx/><c:valAx/>" +
                "</c:plotArea></c:chart></c:chartSpace>");

            using (Bitmap targetLayoutBitmap =
                new Bitmap(
                    360,
                    240,
                    PixelFormat.Format32bppArgb))
            using (Graphics targetLayoutGraphics =
                Graphics.FromImage(
                    targetLayoutBitmap))
            {
                RectangleF innerTargetPlot;
                RectangleF outerTargetPlot;

                PrepareChartSurface(
                    targetLayoutGraphics,
                    new RectangleF(
                        10f,
                        10f,
                        320f,
                        200f),
                    innerTargetChart,
                    theme,
                    out innerTargetPlot);

                PrepareChartSurface(
                    targetLayoutGraphics,
                    new RectangleF(
                        10f,
                        10f,
                        320f,
                        200f),
                    outerTargetChart,
                    theme,
                    out outerTargetPlot);

                if (Math.Abs(
                        innerTargetPlot.Left -
                        42f) > 1.0f ||
                    Math.Abs(
                        innerTargetPlot.Top -
                        30f) > 1.0f ||
                    Math.Abs(
                        innerTargetPlot.Width -
                        192f) > 1.0f ||
                    Math.Abs(
                        innerTargetPlot.Height -
                        120f) > 1.0f)
                {
                    throw new InvalidOperationException(
                        "Chart inner layoutTarget was not applied to the plotting rectangle.");
                }

                if (outerTargetPlot.Left <=
                        innerTargetPlot.Left + 4f ||
                    outerTargetPlot.Right >=
                        innerTargetPlot.Right - 4f ||
                    outerTargetPlot.Bottom >=
                        innerTargetPlot.Bottom - 4f)
                {
                    throw new InvalidOperationException(
                        "Chart outer layoutTarget did not preserve axis-label insets.");
                }
            }

            XmlDocument nestedManualLayoutChart =
                new XmlDocument();

            nestedManualLayoutChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart><c:plotArea>" +
                "<c:barChart><c:dLbls><c:dLbl><c:idx val=\"0\"/><c:layout><c:manualLayout>" +
                "<c:x val=\"0.90\"/><c:y val=\"0.90\"/>" +
                "</c:manualLayout></c:layout></c:dLbl></c:dLbls></c:barChart>" +
                "</c:plotArea></c:chart></c:chartSpace>");

            using (Bitmap nestedLayoutBitmap =
                new Bitmap(
                    360,
                    240,
                    PixelFormat.Format32bppArgb))
            using (Graphics nestedLayoutGraphics =
                Graphics.FromImage(
                    nestedLayoutBitmap))
            {
                RectangleF nestedPlot;

                PrepareChartSurface(
                    nestedLayoutGraphics,
                    new RectangleF(
                        10f,
                        10f,
                        320f,
                        200f),
                    nestedManualLayoutChart,
                    theme,
                    out nestedPlot);

                if (nestedPlot.Left > 60f ||
                    nestedPlot.Top > 45f)
                {
                    throw new InvalidOperationException(
                        "Nested data-label manual layout leaked into plotArea layout scope.");
                }
            }

            XmlDocument elementLayoutChart =
                new XmlDocument();

            elementLayoutChart.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
                "<c:chart>" +
                "<c:title><c:layout><c:manualLayout>" +
                "<c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/>" +
                "<c:x val=\"0.10\"/><c:y val=\"0.08\"/>" +
                "<c:w val=\"0.42\"/><c:h val=\"0.12\"/>" +
                "</c:manualLayout></c:layout></c:title>" +
                "<c:plotArea/>" +
                "<c:legend><c:legendPos val=\"b\"/><c:layout><c:manualLayout>" +
                "<c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/>" +
                "<c:x val=\"0.10\"/><c:y val=\"0.20\"/>" +
                "<c:w val=\"0.40\"/><c:h val=\"0.30\"/>" +
                "</c:manualLayout></c:layout></c:legend>" +
                "</c:chart></c:chartSpace>");

            RectangleF elementChartRect =
                new RectangleF(
                    0f,
                    0f,
                    320f,
                    180f);

            RectangleF resolvedTitleRect;

            if (!TryResolveChartElementManualLayout(
                    elementLayoutChart,
                    "title",
                    elementChartRect,
                    new RectangleF(
                        5f,
                        4f,
                        310f,
                        30f),
                    out resolvedTitleRect) ||
                Math.Abs(
                    resolvedTitleRect.Left -
                    32f) > 0.5f ||
                Math.Abs(
                    resolvedTitleRect.Top -
                    14.4f) > 0.5f ||
                Math.Abs(
                    resolvedTitleRect.Width -
                    134.4f) > 0.5f ||
                Math.Abs(
                    resolvedTitleRect.Height -
                    21.6f) > 0.5f)
            {
                throw new InvalidOperationException(
                    "Chart title manual layout coordinates were not resolved.");
            }

            RectangleF resolvedLegendRect;

            if (!TryResolveChartElementManualLayout(
                    elementLayoutChart,
                    "legend",
                    elementChartRect,
                    new RectangleF(
                        12f,
                        144f,
                        300f,
                        24f),
                    out resolvedLegendRect) ||
                Math.Abs(
                    resolvedLegendRect.Left -
                    32f) > 0.5f ||
                Math.Abs(
                    resolvedLegendRect.Top -
                    36f) > 0.5f ||
                Math.Abs(
                    resolvedLegendRect.Width -
                    128f) > 0.5f ||
                Math.Abs(
                    resolvedLegendRect.Height -
                    54f) > 0.5f)
            {
                throw new InvalidOperationException(
                    "Chart legend manual layout coordinates were not resolved.");
            }

            using (Bitmap manualLegendBitmap =
                new Bitmap(
                    320,
                    180,
                    PixelFormat.Format32bppArgb))
            using (Graphics manualLegendGraphics =
                Graphics.FromImage(
                    manualLegendBitmap))
            {
                manualLegendGraphics.Clear(
                    Color.White);

                List<ChartSeriesData> manualLegendSeries =
                    new List<ChartSeriesData>();

                ChartSeriesData manualLegendItem =
                    new ChartSeriesData();
                manualLegendItem.Name =
                    "Manual";
                manualLegendItem.Values.Add(
                    1.0);
                manualLegendSeries.Add(
                    manualLegendItem);

                Color manualLegendColor =
                    Color.FromArgb(
                        34,
                        116,
                        208);

                DrawChartLegend(
                    manualLegendGraphics,
                    elementChartRect,
                    manualLegendSeries,
                    new Color[]
                    {
                        manualLegendColor
                    },
                    "column",
                    "b",
                    null,
                    null,
                    elementLayoutChart);

                Color manualLegendPixel =
                    manualLegendBitmap.GetPixel(
                        34,
                        40);

                if (Math.Abs(
                        manualLegendPixel.R -
                        manualLegendColor.R) > 3 ||
                    Math.Abs(
                        manualLegendPixel.G -
                        manualLegendColor.G) > 3 ||
                    Math.Abs(
                        manualLegendPixel.B -
                        manualLegendColor.B) > 3)
                {
                    throw new InvalidOperationException(
                        "Chart legend manual layout was not used while rendering.");
                }
            }

            XmlDocument doughnut =
                new XmlDocument();

            doughnut.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:spPr><a:solidFill><a:srgbClr val=\"F0F4FA\"/></a:solidFill></c:spPr>" +
                "<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Doughnut</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
                "<c:plotArea><c:doughnutChart><c:firstSliceAng val=\"75\"/><c:holeSize val=\"66\"/>" +
                "<c:dLbls><c:showCatName val=\"1\"/><c:showPercent val=\"1\"/><c:separator> </c:separator></c:dLbls>" +
                BuildSyntheticChartSeries(
                    0,
                    "Mix",
                    new string[] { "A", "B", "C", "D" },
                    new double[] { 22.0, 31.0, 17.0, 30.0 }) +
                "</c:doughnutChart>" +
                "<c:spPr><a:solidFill><a:srgbClr val=\"EFF7EE\"/></a:solidFill></c:spPr>" +
                "</c:plotArea>" +
                "<c:legend><c:legendPos val=\"b\"/></c:legend></c:chart></c:chartSpace>");

            XmlDocument area =
                new XmlDocument();

            area.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Area</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
                "<c:plotArea><c:areaChart><c:grouping val=\"stacked\"/>" +
                BuildSyntheticChartSeries(
                    0,
                    "North",
                    new string[] { "Jan", "Feb", "Mar", "Apr" },
                    new double[] { 18.0, 35.0, 27.0, 46.0 }) +
                BuildSyntheticChartSeries(
                    1,
                    "South",
                    new string[] { "Jan", "Feb", "Mar", "Apr" },
                    new double[] { 12.0, 23.0, 42.0, 33.0 }) +
                "</c:areaChart>" +
                "<c:catAx><c:axId val=\"11\"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>" +
                "<c:valAx><c:axId val=\"12\"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Total</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>" +
                "</c:plotArea>" +
                "<c:legend><c:legendPos val=\"r\"/></c:legend></c:chart></c:chartSpace>");

            XmlDocument scatter =
                new XmlDocument();

            scatter.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Scatter</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
                "<c:plotArea><c:scatterChart><c:scatterStyle val=\"lineMarker\"/>" +
                "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>Trend</c:v></c:tx>" +
                "<c:spPr><a:solidFill><a:srgbClr val=\"7B4FB3\"/></a:solidFill><a:ln w=\"38100\"><a:prstDash val=\"dash\"/></a:ln></c:spPr>" +
                "<c:marker><c:symbol val=\"diamond\"/><c:size val=\"8\"/></c:marker>" +
                "<c:xVal><c:numLit><c:pt idx=\"0\"><c:v>1</c:v></c:pt><c:pt idx=\"1\"><c:v>2</c:v></c:pt><c:pt idx=\"2\"><c:v>3</c:v></c:pt><c:pt idx=\"3\"><c:v>4</c:v></c:pt></c:numLit></c:xVal>" +
                "<c:yVal><c:numLit><c:pt idx=\"0\"><c:v>1.5</c:v></c:pt><c:pt idx=\"1\"><c:v>3.2</c:v></c:pt><c:pt idx=\"2\"><c:v>2.4</c:v></c:pt><c:pt idx=\"3\"><c:v>4.6</c:v></c:pt></c:numLit></c:yVal>" +
                "</c:ser></c:scatterChart></c:plotArea>" +
                "<c:legend><c:legendPos val=\"t\"/></c:legend></c:chart></c:chartSpace>");

            XmlDocument radar =
                new XmlDocument();

            radar.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Radar</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
                "<c:plotArea><c:radarChart>" +
                BuildSyntheticChartSeries(
                    0,
                    "Score",
                    new string[] { "Speed", "Power", "Range", "Skill", "Guard" },
                    new double[] { 70.0, 82.0, 54.0, 91.0, 66.0 }) +
                "</c:radarChart></c:plotArea></c:chart></c:chartSpace>");

            if (ReadAreaChartGrouping(area) != "stacked")
            {
                throw new InvalidOperationException(
                    "Area chart grouping was not parsed correctly.");
            }

            using (Bitmap bitmap =
                new Bitmap(
                    1200,
                    760,
                    PixelFormat.Format32bppArgb))
            using (Graphics graphics =
                Graphics.FromImage(
                    bitmap))
            {
                graphics.Clear(
                    Color.White);
                graphics.SmoothingMode =
                    SmoothingMode.AntiAlias;

                DrawDoughnutChart(
                    graphics,
                    doughnut,
                    new RectangleF(
                        10f,
                        10f,
                        580f,
                        360f),
                    theme);

                Color extendedChartBackground =
                    bitmap.GetPixel(
                        20,
                        20);
                Color extendedPlotBackground =
                    bitmap.GetPixel(
                        60,
                        60);

                if (Math.Abs(
                        extendedChartBackground.R -
                        0xF0) > 3 ||
                    Math.Abs(
                        extendedChartBackground.G -
                        0xF4) > 3 ||
                    Math.Abs(
                        extendedChartBackground.B -
                        0xFA) > 3 ||
                    Math.Abs(
                        extendedPlotBackground.R -
                        0xEF) > 3 ||
                    Math.Abs(
                        extendedPlotBackground.G -
                        0xF7) > 3 ||
                    Math.Abs(
                        extendedPlotBackground.B -
                        0xEE) > 3)
                {
                    throw new InvalidOperationException(
                        "Extended chart chartSpace or plotArea fill was not rendered.");
                }

                DrawAreaChart(
                    graphics,
                    area,
                    new RectangleF(
                        610f,
                        10f,
                        580f,
                        360f),
                    theme);

                DrawScatterChart(
                    graphics,
                    scatter,
                    new RectangleF(
                        10f,
                        390f,
                        580f,
                        360f),
                    theme,
                    false);

                DrawRadarChart(
                    graphics,
                    radar,
                    new RectangleF(
                        610f,
                        390f,
                        580f,
                        360f),
                    theme);

                int nonWhite;
                int distinct;

                AnalyzeDiagnosticBitmap(
                    bitmap,
                    out nonWhite,
                    out distinct);

                if (nonWhite < 650 ||
                    distinct < 10)
                {
                    throw new InvalidOperationException(
                        "Extended chart rendering did not produce enough visual detail.");
                }

                bitmap.Save(
                    outputPath,
                    ImageFormat.Png);
            }

            RequireDiagnosticFile(
                outputPath);
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
                SyntheticSmartAssistantPoint("A", "Advisor") +
                "</dgm:ptLst>" +
                "<dgm:cxnLst>" +
                SyntheticSmartConnection("1", "2") +
                SyntheticSmartConnection("2", "3") +
                SyntheticSmartConnection("3", "4") +
                SyntheticSmartConnection("4", "5") +
                SyntheticSmartConnection("2", "A") +
                "</dgm:cxnLst>" +
                "</dgm:dataModel>");

            SmartNode hierarchyRoot =
                new SmartNode();
            hierarchyRoot.Id = "root";
            hierarchyRoot.Label = "Root";
            hierarchyRoot.Depth = 0;
            hierarchyRoot.Children.Add(
                "child");
            hierarchyRoot.Children.Add(
                "assistant");

            SmartNode hierarchyChild =
                new SmartNode();
            hierarchyChild.Id = "child";
            hierarchyChild.Label = "Child";
            hierarchyChild.Depth = 1;
            hierarchyChild.ParentId = "root";

            SmartNode hierarchyAssistant =
                new SmartNode();
            hierarchyAssistant.Id = "assistant";
            hierarchyAssistant.Label = "Assistant";
            hierarchyAssistant.Depth = 1;
            hierarchyAssistant.IsAssistant = true;
            hierarchyAssistant.ParentId = "root";

            List<SmartNode> hierarchyNodes =
                new List<SmartNode>();
            hierarchyNodes.Add(
                hierarchyRoot);
            hierarchyNodes.Add(
                hierarchyChild);
            hierarchyNodes.Add(
                hierarchyAssistant);

            Dictionary<string, RectangleF> hierarchyPositions =
                BuildHierarchySmartArtPositions(
                    hierarchyNodes,
                    new RectangleF(
                        0f,
                        0f,
                        420f,
                        260f),
                    1);

            RectangleF rootPosition;
            RectangleF assistantPosition;

            RectangleF childPosition;

            if (!hierarchyPositions.TryGetValue(
                    "root",
                    out rootPosition) ||
                !hierarchyPositions.TryGetValue(
                    "child",
                    out childPosition) ||
                !hierarchyPositions.TryGetValue(
                    "assistant",
                    out assistantPosition) ||
                rootPosition.IntersectsWith(
                    assistantPosition) ||
                (assistantPosition.Left <
                    rootPosition.Right &&
                 assistantPosition.Right >
                    rootPosition.Left))
            {
                throw new InvalidOperationException(
                    "SmartArt hierarchy assistant was not positioned beside its parent.");
            }

            SmartNode multiAssistantRoot =
                new SmartNode();
            multiAssistantRoot.Id =
                "multiAssistantRoot";
            multiAssistantRoot.Label =
                "Root";
            multiAssistantRoot.Depth = 0;

            List<SmartNode> multiAssistantNodes =
                new List<SmartNode>();
            multiAssistantNodes.Add(
                multiAssistantRoot);

            for (int i = 0;
                 i < 4;
                 i++)
            {
                SmartNode assistant =
                    new SmartNode();

                assistant.Id =
                    "multiAssistant" +
                    i.ToString();
                assistant.Label =
                    "Assistant " +
                    (i + 1).ToString();
                assistant.Depth = 1;
                assistant.IsAssistant = true;
                assistant.ParentId =
                    multiAssistantRoot.Id;

                multiAssistantNodes.Add(
                    assistant);
            }

            Dictionary<string, RectangleF> multiAssistantPositions =
                BuildHierarchySmartArtPositions(
                    multiAssistantNodes,
                    new RectangleF(
                        0f,
                        0f,
                        900f,
                        360f),
                    1);

            RectangleF multiRootBox =
                multiAssistantPositions[
                    "multiAssistantRoot"];
            RectangleF assistant0 =
                multiAssistantPositions[
                    "multiAssistant0"];
            RectangleF assistant1 =
                multiAssistantPositions[
                    "multiAssistant1"];
            RectangleF assistant2 =
                multiAssistantPositions[
                    "multiAssistant2"];
            RectangleF assistant3 =
                multiAssistantPositions[
                    "multiAssistant3"];

            float multiParentCenterY =
                multiRootBox.Top +
                multiRootBox.Height /
                2f;
            float leftStackCenterY =
                (assistant0.Top +
                 assistant0.Height /
                 2f +
                 assistant2.Top +
                 assistant2.Height /
                 2f) /
                2f;
            float rightStackCenterY =
                (assistant1.Top +
                 assistant1.Height /
                 2f +
                 assistant3.Top +
                 assistant3.Height /
                 2f) /
                2f;

            if (assistant0.Right >
                    multiRootBox.Left ||
                assistant2.Right >
                    multiRootBox.Left ||
                assistant1.Left <
                    multiRootBox.Right ||
                assistant3.Left <
                    multiRootBox.Right ||
                assistant0.Bottom >
                    assistant2.Top ||
                assistant1.Bottom >
                    assistant3.Top ||
                Math.Abs(
                    leftStackCenterY -
                    multiParentCenterY) > 2f ||
                Math.Abs(
                    rightStackCenterY -
                    multiParentCenterY) > 2f)
            {
                throw new InvalidOperationException(
                    "SmartArt multiple assistants were not separated into centered left/right stacks.");
            }

            SmartNode leftParent =
                new SmartNode();
            leftParent.Id = "leftParent";
            leftParent.Label = "Left";
            leftParent.Depth = 0;
            leftParent.Children.Add(
                "leftChild1");
            leftParent.Children.Add(
                "leftChild2");

            SmartNode rightParent =
                new SmartNode();
            rightParent.Id = "rightParent";
            rightParent.Label = "Right";
            rightParent.Depth = 0;
            rightParent.Children.Add(
                "rightChild1");
            rightParent.Children.Add(
                "rightChild2");

            SmartNode rightChild1 =
                new SmartNode();
            rightChild1.Id = "rightChild1";
            rightChild1.Depth = 1;
            rightChild1.ParentId =
                "rightParent";

            SmartNode leftChild1 =
                new SmartNode();
            leftChild1.Id = "leftChild1";
            leftChild1.Depth = 1;
            leftChild1.ParentId =
                "leftParent";

            SmartNode rightChild2 =
                new SmartNode();
            rightChild2.Id = "rightChild2";
            rightChild2.Depth = 1;
            rightChild2.ParentId =
                "rightParent";

            SmartNode leftChild2 =
                new SmartNode();
            leftChild2.Id = "leftChild2";
            leftChild2.Depth = 1;
            leftChild2.ParentId =
                "leftParent";

            List<SmartNode> groupedHierarchyNodes =
                new List<SmartNode>();
            groupedHierarchyNodes.Add(
                leftParent);
            groupedHierarchyNodes.Add(
                rightParent);
            groupedHierarchyNodes.Add(
                rightChild1);
            groupedHierarchyNodes.Add(
                leftChild1);
            groupedHierarchyNodes.Add(
                rightChild2);
            groupedHierarchyNodes.Add(
                leftChild2);

            Dictionary<string, RectangleF> groupedPositions =
                BuildHierarchySmartArtPositions(
                    groupedHierarchyNodes,
                    new RectangleF(
                        0f,
                        0f,
                        600f,
                        300f),
                    1);

            RectangleF leftParentBox =
                groupedPositions[
                    "leftParent"];
            RectangleF rightParentBox =
                groupedPositions[
                    "rightParent"];
            RectangleF leftChildBox1 =
                groupedPositions[
                    "leftChild1"];
            RectangleF leftChildBox2 =
                groupedPositions[
                    "leftChild2"];
            RectangleF rightChildBox1 =
                groupedPositions[
                    "rightChild1"];
            RectangleF rightChildBox2 =
                groupedPositions[
                    "rightChild2"];

            float leftParentCenter =
                leftParentBox.Left +
                leftParentBox.Width /
                2f;
            float rightParentCenter =
                rightParentBox.Left +
                rightParentBox.Width /
                2f;
            float leftChildrenCenter =
                (leftChildBox1.Left +
                 leftChildBox1.Width /
                 2f +
                 leftChildBox2.Left +
                 leftChildBox2.Width /
                 2f) /
                2f;
            float rightChildrenCenter =
                (rightChildBox1.Left +
                 rightChildBox1.Width /
                 2f +
                 rightChildBox2.Left +
                 rightChildBox2.Width /
                 2f) /
                2f;

            if (Math.Abs(
                    leftChildrenCenter -
                    leftParentCenter) >=
                Math.Abs(
                    leftChildrenCenter -
                    rightParentCenter) ||
                Math.Abs(
                    rightChildrenCenter -
                    rightParentCenter) >=
                Math.Abs(
                    rightChildrenCenter -
                    leftParentCenter) ||
                Math.Max(
                    leftChildBox1.Right,
                    leftChildBox2.Right) >
                Math.Min(
                    rightChildBox1.Left,
                    rightChildBox2.Left))
            {
                throw new InvalidOperationException(
                    "SmartArt hierarchy children were not grouped beneath their parent nodes.");
            }

            Dictionary<string, SmartNode> hierarchyNodeMap =
                new Dictionary<string, SmartNode>(
                    StringComparer.Ordinal);
            hierarchyNodeMap[
                hierarchyRoot.Id] =
                hierarchyRoot;
            hierarchyNodeMap[
                hierarchyChild.Id] =
                hierarchyChild;
            hierarchyNodeMap[
                hierarchyAssistant.Id] =
                hierarchyAssistant;

            using (Bitmap connectorBitmap =
                new Bitmap(
                    420,
                    260,
                    PixelFormat.Format32bppArgb))
            using (Graphics connectorGraphics =
                Graphics.FromImage(
                    connectorBitmap))
            {
                connectorGraphics.Clear(
                    Color.White);

                DrawSmartArtConnectors(
                    connectorGraphics,
                    "hierarchy",
                    hierarchyNodes,
                    hierarchyNodeMap,
                    hierarchyPositions,
                    new Color[]
                    {
                        Color.FromArgb(
                            76,
                            123,
                            205)
                    });

                float regularMidY =
                    rootPosition.Bottom +
                    (childPosition.Top -
                     rootPosition.Bottom) /
                    2f;

                Color verticalElbowPixel =
                    connectorBitmap.GetPixel(
                        Math.Max(
                            0,
                            Math.Min(
                                connectorBitmap.Width - 1,
                                (int)Math.Round(
                                    rootPosition.Left +
                                    rootPosition.Width /
                                    2f))),
                        Math.Max(
                            0,
                            Math.Min(
                                connectorBitmap.Height - 1,
                                (int)Math.Round(
                                    regularMidY))));

                bool assistantOnLeft =
                    assistantPosition.Left <
                    rootPosition.Left;
                float assistantFromX =
                    assistantOnLeft
                        ? rootPosition.Left
                        : rootPosition.Right;
                float assistantToX =
                    assistantOnLeft
                        ? assistantPosition.Right
                        : assistantPosition.Left;
                float assistantMidX =
                    assistantFromX +
                    (assistantToX -
                     assistantFromX) /
                    2f;

                Color assistantElbowPixel =
                    connectorBitmap.GetPixel(
                        Math.Max(
                            0,
                            Math.Min(
                                connectorBitmap.Width - 1,
                                (int)Math.Round(
                                    assistantMidX))),
                        Math.Max(
                            0,
                            Math.Min(
                                connectorBitmap.Height - 1,
                                (int)Math.Round(
                                    rootPosition.Top +
                                    rootPosition.Height /
                                    2f))));

                bool regularConnectorVisible =
                    verticalElbowPixel.R <
                        245 ||
                    verticalElbowPixel.G <
                        245 ||
                    verticalElbowPixel.B <
                        245;
                bool assistantConnectorVisible =
                    assistantElbowPixel.R <
                        245 ||
                    assistantElbowPixel.G <
                        245 ||
                    assistantElbowPixel.B <
                        245;

                if (!regularConnectorVisible ||
                    !assistantConnectorVisible)
                {
                    throw new InvalidOperationException(
                        "SmartArt hierarchy elbow connector routing was not rendered.");
                }
            }

            SmartNode processLayoutNode0 =
                new SmartNode();
            processLayoutNode0.Id = "processLayout0";
            processLayoutNode0.Label = "P1";

            SmartNode processLayoutNode1 =
                new SmartNode();
            processLayoutNode1.Id = "processLayout1";
            processLayoutNode1.Label = "P2";

            SmartNode processLayoutNode2 =
                new SmartNode();
            processLayoutNode2.Id = "processLayout2";
            processLayoutNode2.Label = "P3";

            SmartNode processLayoutNode3 =
                new SmartNode();
            processLayoutNode3.Id = "processLayout3";
            processLayoutNode3.Label = "P4";

            SmartNode processLayoutNode4 =
                new SmartNode();
            processLayoutNode4.Id = "processLayout4";
            processLayoutNode4.Label = "P5";

            SmartNode processLayoutNode5 =
                new SmartNode();
            processLayoutNode5.Id = "processLayout5";
            processLayoutNode5.Label = "P6";

            List<SmartNode> processLayoutNodes =
                new List<SmartNode>();
            processLayoutNodes.Add(
                processLayoutNode0);
            processLayoutNodes.Add(
                processLayoutNode1);
            processLayoutNodes.Add(
                processLayoutNode2);
            processLayoutNodes.Add(
                processLayoutNode3);
            processLayoutNodes.Add(
                processLayoutNode4);
            processLayoutNodes.Add(
                processLayoutNode5);

            RectangleF processLayoutRect =
                new RectangleF(
                    0f,
                    0f,
                    600f,
                    300f);

            Dictionary<string, RectangleF> processLayoutPositions =
                BuildProcessSmartArtPositions(
                    processLayoutNodes,
                    processLayoutRect);

            RectangleF processFirstRow =
                processLayoutPositions[
                    "processLayout0"];
            RectangleF processLastRow =
                processLayoutPositions[
                    "processLayout5"];

            float lastRowCenter =
                processLastRow.Left +
                processLastRow.Width /
                2f;
            float expectedProcessCenter =
                processLayoutRect.Left +
                processLayoutRect.Width /
                2f;

            if (Math.Abs(
                    lastRowCenter -
                    expectedProcessCenter) > 1.5f ||
                processLastRow.Top <=
                    processFirstRow.Bottom ||
                processLastRow.Left <
                    processLayoutRect.Left ||
                processLastRow.Right >
                    processLayoutRect.Right)
            {
                throw new InvalidOperationException(
                    "SmartArt process rows were not centered or spaced correctly.");
            }

            RectangleF narrowProcessRect =
                new RectangleF(
                    0f,
                    0f,
                    180f,
                    300f);

            Dictionary<string, RectangleF> narrowProcessPositions =
                BuildProcessSmartArtPositions(
                    processLayoutNodes,
                    narrowProcessRect);

            for (int processIndex = 0;
                 processIndex < processLayoutNodes.Count;
                 processIndex++)
            {
                RectangleF narrowBox =
                    narrowProcessPositions[
                        processLayoutNodes[
                            processIndex].Id];

                if (narrowBox.Left <
                        narrowProcessRect.Left - 0.5f ||
                    narrowBox.Right >
                        narrowProcessRect.Right + 0.5f)
                {
                    throw new InvalidOperationException(
                        "SmartArt process layout overflowed a narrow container.");
                }
            }

            if (narrowProcessPositions[
                    "processLayout3"].Top <=
                narrowProcessPositions[
                    "processLayout0"].Bottom)
            {
                throw new InvalidOperationException(
                    "SmartArt process layout did not reduce columns for a narrow container.");
            }

            SmartNode verticalLayoutNode0 =
                new SmartNode();
            verticalLayoutNode0.Id =
                "verticalLayout0";
            SmartNode verticalLayoutNode1 =
                new SmartNode();
            verticalLayoutNode1.Id =
                "verticalLayout1";
            SmartNode verticalLayoutNode2 =
                new SmartNode();
            verticalLayoutNode2.Id =
                "verticalLayout2";

            List<SmartNode> verticalLayoutNodes =
                new List<SmartNode>();
            verticalLayoutNodes.Add(
                verticalLayoutNode0);
            verticalLayoutNodes.Add(
                verticalLayoutNode1);
            verticalLayoutNodes.Add(
                verticalLayoutNode2);

            RectangleF verticalLayoutRect =
                new RectangleF(
                    0f,
                    0f,
                    320f,
                    520f);

            Dictionary<string, RectangleF> verticalLayoutPositions =
                BuildVerticalProcessSmartArtPositions(
                    verticalLayoutNodes,
                    verticalLayoutRect);

            RectangleF verticalFirst =
                verticalLayoutPositions[
                    "verticalLayout0"];
            RectangleF verticalMiddle =
                verticalLayoutPositions[
                    "verticalLayout1"];
            RectangleF verticalLast =
                verticalLayoutPositions[
                    "verticalLayout2"];

            float topMargin =
                verticalFirst.Top -
                verticalLayoutRect.Top;
            float bottomMargin =
                verticalLayoutRect.Bottom -
                verticalLast.Bottom;
            float verticalCenter =
                verticalLayoutRect.Left +
                verticalLayoutRect.Width /
                2f;
            float firstCenter =
                verticalFirst.Left +
                verticalFirst.Width /
                2f;

            if (Math.Abs(
                    topMargin -
                    bottomMargin) > 2f ||
                Math.Abs(
                    firstCenter -
                    verticalCenter) > 1f ||
                verticalMiddle.Top <=
                    verticalFirst.Bottom ||
                verticalLast.Top <=
                    verticalMiddle.Bottom)
            {
                throw new InvalidOperationException(
                    "SmartArt vertical process stack was not centered or spaced correctly.");
            }

            SmartNode processSource =
                new SmartNode();
            processSource.Id = "processSource";
            processSource.Children.Add(
                "processSameRow");
            processSource.Children.Add(
                "processNextRow");

            SmartNode processSameRow =
                new SmartNode();
            processSameRow.Id = "processSameRow";

            SmartNode processNextRow =
                new SmartNode();
            processNextRow.Id = "processNextRow";

            List<SmartNode> processNodes =
                new List<SmartNode>();
            processNodes.Add(
                processSource);
            processNodes.Add(
                processSameRow);
            processNodes.Add(
                processNextRow);

            Dictionary<string, SmartNode> processNodeMap =
                new Dictionary<string, SmartNode>(
                    StringComparer.Ordinal);
            processNodeMap[
                processSource.Id] =
                processSource;
            processNodeMap[
                processSameRow.Id] =
                processSameRow;
            processNodeMap[
                processNextRow.Id] =
                processNextRow;

            Dictionary<string, RectangleF> processPositions =
                new Dictionary<string, RectangleF>(
                    StringComparer.Ordinal);
            processPositions[
                processSource.Id] =
                new RectangleF(
                    20f,
                    20f,
                    80f,
                    40f);
            processPositions[
                processSameRow.Id] =
                new RectangleF(
                    140f,
                    20f,
                    80f,
                    40f);
            processPositions[
                processNextRow.Id] =
                new RectangleF(
                    40f,
                    120f,
                    80f,
                    40f);

            using (Bitmap processConnectorBitmap =
                new Bitmap(
                    260,
                    200,
                    PixelFormat.Format32bppArgb))
            using (Graphics processConnectorGraphics =
                Graphics.FromImage(
                    processConnectorBitmap))
            {
                processConnectorGraphics.Clear(
                    Color.White);

                DrawSmartArtConnectors(
                    processConnectorGraphics,
                    "process",
                    processNodes,
                    processNodeMap,
                    processPositions,
                    new Color[]
                    {
                        Color.FromArgb(
                            76,
                            123,
                            205)
                    });

                Color sameRowPixel =
                    processConnectorBitmap.GetPixel(
                        120,
                        40);
                Color nextRowElbowPixel =
                    processConnectorBitmap.GetPixel(
                        70,
                        90);

                bool sameRowVisible =
                    sameRowPixel.R < 245 ||
                    sameRowPixel.G < 245 ||
                    sameRowPixel.B < 245;
                bool elbowVisible =
                    nextRowElbowPixel.R < 245 ||
                    nextRowElbowPixel.G < 245 ||
                    nextRowElbowPixel.B < 245;

                if (!sameRowVisible ||
                    !elbowVisible)
                {
                    throw new InvalidOperationException(
                        "SmartArt process edge/elbow connector routing was not rendered.");
                }

                processConnectorGraphics.Clear(
                    Color.White);

                processSource.Children.Clear();
                processSource.Children.Add(
                    processNextRow.Id);

                processPositions[
                    processNextRow.Id] =
                    new RectangleF(
                        180f,
                        120f,
                        80f,
                        40f);

                processPositions[
                    "processBlocker"] =
                    new RectangleF(
                        110f,
                        70f,
                        50f,
                        40f);

                DrawSmartArtConnectors(
                    processConnectorGraphics,
                    "process",
                    processNodes,
                    processNodeMap,
                    processPositions,
                    new Color[]
                    {
                        Color.FromArgb(
                            76,
                            123,
                            205)
                    });

                Color blockedMiddlePixel =
                    processConnectorBitmap.GetPixel(
                        130,
                        90);

                Color clearChannelPixel =
                    processConnectorBitmap.GetPixel(
                        130,
                        64);

                bool blockedMiddleStayedClear =
                    blockedMiddlePixel.R >= 245 &&
                    blockedMiddlePixel.G >= 245 &&
                    blockedMiddlePixel.B >= 245;

                bool clearChannelVisible =
                    clearChannelPixel.R < 245 ||
                    clearChannelPixel.G < 245 ||
                    clearChannelPixel.B < 245;

                if (!blockedMiddleStayedClear ||
                    !clearChannelVisible)
                {
                    throw new InvalidOperationException(
                        "SmartArt process connector did not avoid an intervening node.");
                }

                processPositions.Remove(
                    "processBlocker");

                processConnectorGraphics.Clear(
                    Color.White);

                Dictionary<string, RectangleF> verticalPositions =
                    new Dictionary<string, RectangleF>(
                        StringComparer.Ordinal);
                verticalPositions[
                    processSource.Id] =
                    new RectangleF(
                        20f,
                        20f,
                        80f,
                        40f);
                verticalPositions[
                    processSameRow.Id] =
                    new RectangleF(
                        140f,
                        120f,
                        80f,
                        40f);

                processSource.Children.Clear();
                processSource.Children.Add(
                    processSameRow.Id);

                DrawSmartArtConnectors(
                    processConnectorGraphics,
                    "verticalProcess",
                    processNodes,
                    processNodeMap,
                    verticalPositions,
                    new Color[]
                    {
                        Color.FromArgb(
                            76,
                            123,
                            205)
                    });

                Color verticalElbow =
                    processConnectorBitmap.GetPixel(
                        120,
                        90);

                if (verticalElbow.R >= 245 &&
                    verticalElbow.G >= 245 &&
                    verticalElbow.B >= 245)
                {
                    throw new InvalidOperationException(
                        "SmartArt vertical-process elbow routing was not rendered.");
                }
            }

            string[] kinds =
                new string[]
                {
                    "Hierarchy",
                    "Basic Process",
                    "Vertical Process",
                    "Cycle",
                    "Matrix",
                    "Pyramid",
                    "Radial Relationship",
                    "Basic Venn"
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
                    1600,
                    900,
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
                        i % 4;
                    int row =
                        i / 4;

                    RectangleF rect =
                        new RectangleF(
                            12f +
                                column *
                                397f,
                            12f +
                                row *
                                438f,
                            385f,
                            426f);

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

                if (nonWhite < 650 ||
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

        private static string SyntheticSmartAssistantPoint(
            string id,
            string label)
        {
            return
                "<dgm:pt modelId=\"" +
                System.Security.SecurityElement.Escape(id) +
                "\" type=\"asst\"><dgm:t><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>" +
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
