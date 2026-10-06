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
                "<c:title><c:tx><c:rich><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Quarterly Delta</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
                "<c:plotArea><c:barChart><c:barDir val=\"col\"/>" +
                "<c:dLbls><c:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"1\"/><c:separator> · </c:separator></c:dLbls>" +
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
                "<c:catAx><c:axId val=\"1\"/><c:tickLblPos val=\"none\"/><c:tickLblSkip val=\"2\"/><c:crossesAt val=\"20\"/><c:spPr><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"25400\"><a:solidFill><a:srgbClr val=\"3366CC\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln></c:spPr><c:title><c:tx><c:rich><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>" +
                "<c:valAx><c:axId val=\"2\"/><c:tickLblPos val=\"none\"/><c:scaling><c:orientation val=\"minMax\"/><c:min val=\"-50\"/><c:max val=\"80\"/></c:scaling><c:majorUnit val=\"20\"/><c:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/><c:majorGridlines><c:spPr><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"12700\"><a:solidFill><a:srgbClr val=\"88AACC\"/></a:solidFill><a:prstDash val=\"dashDot\"/></a:ln></c:spPr></c:majorGridlines><c:spPr><a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" w=\"19050\"><a:solidFill><a:srgbClr val=\"CC5533\"/></a:solidFill><a:prstDash val=\"dot\"/></a:ln></c:spPr><c:title><c:tx><c:rich><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:t>Delta</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>" +
                "<c:spPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:srgbClr val=\"FFF8EE\"/></a:solidFill></c:spPr>" +
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
                    1);

                bool foundAbovePlot =
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
                            break;
                        }
                    }
                }

                if (!foundAbovePlot)
                {
                    throw new InvalidOperationException(
                        "Chart category tickLblPos=high did not move labels above the plot.");
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

            ChartLabelOptions labels =
                ReadChartLabelOptions(
                    chart);

            if (!labels.ShowValue ||
                !labels.ShowCategoryName ||
                labels.ShowSeriesName ||
                labels.ShowPercent ||
                labels.NumberFormat != "0.0" ||
                labels.Separator != " · ")
            {
                throw new InvalidOperationException(
                    "Chart data-label options were not parsed correctly.");
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

            if (ReadChartLegendPosition(chart) != "b")
            {
                throw new InvalidOperationException(
                    "Chart legend position was not parsed correctly.");
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

            XmlDocument doughnut =
                new XmlDocument();

            doughnut.LoadXml(
                "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Doughnut</a:t></a:r></a:p></c:rich></c:tx></c:title>" +
                "<c:plotArea><c:doughnutChart><c:firstSliceAng val=\"75\"/><c:holeSize val=\"66\"/>" +
                "<c:dLbls><c:showCatName val=\"1\"/><c:showPercent val=\"1\"/><c:separator> </c:separator></c:dLbls>" +
                BuildSyntheticChartSeries(
                    0,
                    "Mix",
                    new string[] { "A", "B", "C", "D" },
                    new double[] { 22.0, 31.0, 17.0, 30.0 }) +
                "</c:doughnutChart></c:plotArea>" +
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
