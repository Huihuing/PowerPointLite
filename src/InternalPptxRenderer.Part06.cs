using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace PptxViewer
{
internal static partial class InternalPptxRenderer
    {
        private sealed class ChartAxisScale
        {
            public double Minimum;
            public double Maximum;
            public double MajorUnit;
            public double MinorUnit;
            public bool Reverse;
            public string NumberFormat;
        }

        private sealed class ChartLineStyle
        {
            public Color Color =
                Color.FromArgb(
                    100,
                    100,
                    100);
            public float Width = 1.2f;
            public DashStyle DashStyle =
                DashStyle.Solid;
        }

        private sealed class ChartLabelOptions
        {
            public bool ShowValue;
            public bool ShowCategoryName;
            public bool ShowSeriesName;
            public bool ShowPercent;
            public bool ShowLeaderLines;
            public Color? TextColor;
            public string NumberFormat;
            public string Position;
            public string Separator = ", ";

            public bool HasAny
            {
                get
                {
                    return ShowValue ||
                        ShowCategoryName ||
                        ShowSeriesName ||
                        ShowPercent;
                }
            }
        }

        private sealed class ChartPointLabelOverride
        {
            public bool? ShowValue;
            public bool? ShowCategoryName;
            public bool? ShowSeriesName;
            public bool? ShowPercent;
            public bool Delete;
            public Color? TextColor;
            public string NumberFormat;
            public string Position;
            public string Separator;
        }

        private sealed class ChartBarOptions
        {
            public string Grouping = "clustered";
            public float GapWidth = 150f;
            public float Overlap;

            public bool IsStacked
            {
                get
                {
                    return string.Equals(
                        Grouping,
                        "stacked",
                        StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(
                            Grouping,
                            "percentStacked",
                            StringComparison.OrdinalIgnoreCase);
                }
            }

            public bool IsPercentStacked
            {
                get
                {
                    return string.Equals(
                        Grouping,
                        "percentStacked",
                        StringComparison.OrdinalIgnoreCase);
                }
            }
        }

        private static void DrawChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (chartDoc == null)
                return;

            string kind = "column";

            if (FindFirst(chartDoc, "pieChart") != null ||
                FindFirst(chartDoc, "doughnutChart") != null)
            {
                kind = "pie";
            }
            else if (FindFirst(chartDoc, "lineChart") != null)
            {
                kind = "line";
            }
            else if (FindFirst(chartDoc, "barChart") != null)
            {
                XmlNode barDir = FindFirst(chartDoc, "barDir");
                kind = barDir != null &&
                    GetAttr(barDir, "val") == "bar"
                        ? "bar"
                        : "column";
            }
            else if (FindFirst(chartDoc, "areaChart") != null)
            {
                kind = "line";
            }

            List<ChartSeriesData> series =
                new List<ChartSeriesData>();

            foreach (XmlNode ser in FindAll(chartDoc, "ser"))
            {
                ChartSeriesData data =
                    new ChartSeriesData();

                data.Name =
                    ReadChartSeriesName(ser);
                ReadChartCategories(
                    ser,
                    data.Categories);
                ReadChartValues(
                    ser,
                    data.Values);

                data.ExplicitColor =
                    ReadChartSeriesColor(
                        ser,
                        theme);

                ReadChartPointColors(
                    ser,
                    theme,
                    data.PointColors,
                    data.Values.Count);

                ReadChartSeriesVisualStyle(
                    ser,
                    data);

                ReadChartSeriesLabelOverrides(
                    ser,
                    data,
                    theme);

                if (data.Values.Count > 0)
                    series.Add(data);
            }

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
                g.FillRectangle(bg, rect);
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

            if (series.Count == 0)
            {
                DrawPlaceholder(
                    g,
                    rect,
                    "Chart (no cached data)");
                return;
            }

            string title =
                ReadChartTitle(chartDoc);

            float topPad =
                string.IsNullOrEmpty(title)
                    ? 12f
                    : Math.Max(
                        34f,
                        rect.Height * 0.10f);

            if (!string.IsNullOrEmpty(title))
            {
                using (Font font = SafeFont(
                    "Arial",
                    Math.Max(
                        10f,
                        Math.Min(
                            18f,
                            rect.Height / 18f)),
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
                            rect.Left + 5,
                            rect.Top + 4,
                            rect.Width - 10,
                            topPad - 4),
                        sf);
                }
            }

            string legendPosition =
                ReadChartLegendPosition(
                    chartDoc);

            string valueAxisTitle =
                ReadChartAxisTitle(
                    chartDoc,
                    "valAx");

            string categoryAxisTitle =
                ReadChartAxisTitle(
                    chartDoc,
                    "catAx");

            float leftPad =
                Math.Max(
                    36f,
                    rect.Width * 0.09f);
            float rightPad =
                Math.Max(
                    28f,
                    rect.Width * 0.06f);
            float bottomPad =
                Math.Max(
                    36f,
                    rect.Height * 0.11f);
            float plotTop =
                rect.Top +
                topPad;

            if (legendPosition == "r" ||
                legendPosition == "tr")
            {
                rightPad =
                    Math.Max(
                        92f,
                        rect.Width * 0.20f);
            }
            else if (legendPosition == "l")
            {
                leftPad =
                    Math.Max(
                        92f,
                        rect.Width * 0.20f);
            }
            else if (legendPosition == "b")
            {
                bottomPad =
                    Math.Max(
                        66f,
                        rect.Height * 0.18f);
            }
            else if (legendPosition == "t")
            {
                plotTop +=
                    Math.Max(
                        28f,
                        rect.Height * 0.10f);
            }

            if (kind == "bar")
            {
                if (!string.IsNullOrEmpty(
                        valueAxisTitle))
                {
                    bottomPad += 24f;
                }

                if (!string.IsNullOrEmpty(
                        categoryAxisTitle))
                {
                    leftPad += 24f;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(
                        valueAxisTitle))
                {
                    leftPad += 26f;
                }

                if (!string.IsNullOrEmpty(
                        categoryAxisTitle))
                {
                    bottomPad += 24f;
                }
            }

            RectangleF plot =
                new RectangleF(
                    rect.Left +
                        leftPad,
                    plotTop,
                    Math.Max(
                        10f,
                        rect.Width -
                            leftPad -
                            rightPad),
                    Math.Max(
                        10f,
                        rect.Bottom -
                            plotTop -
                            bottomPad));

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

            Color[] palette =
                new Color[]
                {
                    ThemeOrDefault(
                        theme,
                        "accent1",
                        Color.FromArgb(
                            79,
                            129,
                            189)),
                    ThemeOrDefault(
                        theme,
                        "accent2",
                        Color.FromArgb(
                            192,
                            80,
                            77)),
                    ThemeOrDefault(
                        theme,
                        "accent3",
                        Color.FromArgb(
                            155,
                            187,
                            89)),
                    ThemeOrDefault(
                        theme,
                        "accent4",
                        Color.FromArgb(
                            128,
                            100,
                            162)),
                    ThemeOrDefault(
                        theme,
                        "accent5",
                        Color.FromArgb(
                            75,
                            172,
                            198)),
                    ThemeOrDefault(
                        theme,
                        "accent6",
                        Color.FromArgb(
                            247,
                            150,
                            70))
                };

            ChartLabelOptions labelOptions =
                ReadChartLabelOptions(
                    chartDoc,
                    theme);

            ChartBarOptions barOptions =
                ReadChartBarOptions(
                    chartDoc);

            if (kind == "pie")
            {
                float firstSliceAngle =
                    ReadChartFirstSliceAngle(
                        chartDoc);

                DrawPieChart(
                    g,
                    plot,
                    series[0],
                    palette,
                    firstSliceAngle);

                if (HasAnyChartSeriesDataLabel(
                        labelOptions,
                        series[0]))
                {
                    DrawPieChartValueLabels(
                        g,
                        plot,
                        series[0],
                        labelOptions,
                        firstSliceAngle,
                        ReadChartLeaderLineStyle(
                            chartDoc,
                            theme));
                }

                DrawChartLegend(
                    g,
                    rect,
                    series,
                    palette,
                    kind,
                    legendPosition);
                return;
            }

            double minValue = 0.0;
            double maxValue = 0.0;
            int categoryCount = 0;

            for (int sIndex = 0;
                 sIndex < series.Count;
                 sIndex++)
            {
                categoryCount =
                    Math.Max(
                        categoryCount,
                        series[sIndex].Values.Count);

                for (int i = 0;
                     i < series[sIndex].Values.Count;
                     i++)
                {
                    double value =
                        series[sIndex].Values[i];

                    minValue =
                        Math.Min(
                            minValue,
                            value);
                    maxValue =
                        Math.Max(
                            maxValue,
                            value);
                }
            }

            if (categoryCount <= 0)
                categoryCount = 1;

            if ((kind == "bar" ||
                 kind == "column") &&
                barOptions.IsStacked)
            {
                CalculateStackedChartRange(
                    series,
                    categoryCount,
                    barOptions.IsPercentStacked,
                    out minValue,
                    out maxValue);
            }

            if (Math.Abs(
                    maxValue -
                    minValue) < 0.0000001)
            {
                maxValue =
                    minValue + 1.0;
            }

            ChartAxisScale axisScale =
                ReadChartAxisScale(
                    chartDoc,
                    minValue,
                    maxValue);

            if (barOptions.IsPercentStacked &&
                string.IsNullOrEmpty(
                    axisScale.NumberFormat))
            {
                axisScale.NumberFormat =
                    "0%";
            }

            minValue =
                axisScale.Minimum;
            maxValue =
                axisScale.Maximum;

            string valueTickLabelPosition =
                ReadChartAxisTickLabelPosition(
                    chartDoc,
                    "valAx");

            DrawChartValueGrid(
                g,
                plot,
                axisScale,
                kind,
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
                    theme));

            double range =
                Math.Max(
                    0.0000001,
                    maxValue - minValue);

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

            float zeroX =
                plot.Left +
                (float)(
                    ChartAxisFraction(
                        categoryAxisCrossValue,
                        axisScale) *
                    plot.Width);

            if (kind == "bar")
            {
                zeroX =
                    Math.Max(
                        plot.Left,
                        Math.Min(
                            plot.Right,
                            zeroX));
            }
            else
            {
                zeroY =
                    Math.Max(
                        plot.Top,
                        Math.Min(
                            plot.Bottom,
                            zeroY));
            }

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

                if (kind == "bar")
                {
                    g.DrawLine(
                        categoryAxis,
                        zeroX,
                        plot.Top,
                        zeroX,
                        plot.Bottom);

                    g.DrawLine(
                        valueAxis,
                        plot.Left,
                        plot.Bottom,
                        plot.Right,
                        plot.Bottom);
                }
                else
                {
                    g.DrawLine(
                        categoryAxis,
                        plot.Left,
                        zeroY,
                        plot.Right,
                        zeroY);

                    g.DrawLine(
                        valueAxis,
                        plot.Left,
                        plot.Top,
                        plot.Left,
                        plot.Bottom);
                }
            }

            using (Font valueFont = SafeFont(
                "Arial",
                Math.Max(
                    6f,
                    Math.Min(
                        9f,
                        plot.Height / 36f))))
            using (Brush valueBrush =
                new SolidBrush(
                    Color.FromArgb(
                        75,
                        75,
                        75)))
            {
                if (kind == "line")
                {
                    for (int si = 0;
                         si < series.Count;
                         si++)
                    {
                        ChartSeriesData sd =
                            series[si];

                        if (sd.Values.Count == 0)
                            continue;

                        List<PointF> points =
                            new List<PointF>();

                        for (int i = 0;
                             i < sd.Values.Count;
                             i++)
                        {
                            float x =
                                categoryCount <= 1
                                    ? plot.Left +
                                        plot.Width / 2f
                                    : plot.Left +
                                        plot.Width *
                                        i /
                                        (categoryCount - 1f);

                            float y =
                                plot.Bottom -
                                (float)(
                                    ChartAxisFraction(
                                        sd.Values[i],
                                        axisScale) *
                                    plot.Height);

                            points.Add(
                                new PointF(
                                    x,
                                    y));
                        }

                        Color color =
                            GetChartSeriesColor(
                                series,
                                si,
                                palette);

                        using (Pen pen =
                            new Pen(
                                color,
                                Math.Max(
                                    1f,
                                    sd.LineWidth)))
                        {
                            pen.DashStyle =
                                sd.LineDashStyle;

                            if (points.Count > 1)
                            {
                                g.DrawLines(
                                    pen,
                                    points.ToArray());
                            }

                            for (int i = 0;
                                 i < points.Count;
                                 i++)
                            {
                                PointF point =
                                    points[i];

                                if (sd.MarkerEnabled)
                                {
                                    DrawChartMarker(
                                        g,
                                        point,
                                        sd.MarkerSymbol,
                                        sd.MarkerSize,
                                        color);
                                }

                                if (i < sd.Values.Count)
                                {
                                    ChartLabelOptions pointLabels =
                                        ResolveChartPointLabelOptions(
                                            labelOptions,
                                            sd,
                                            i);

                                    if (pointLabels.HasAny)
                                    {
                                        string label =
                                            BuildChartDataLabel(
                                                pointLabels,
                                                sd,
                                                i,
                                                false);

                                        DrawChartPointDataLabel(
                                            g,
                                            valueFont,
                                            valueBrush,
                                            label,
                                            point,
                                            pointLabels.Position);
                                    }
                                }
                            }
                        }
                    }
                }
                else if (kind == "bar")
                {
                    float groupH =
                        plot.Height /
                        categoryCount;

                    float barH;
                    float barStep;
                    float barSpan;

                    CalculateChartBarLayout(
                        groupH,
                        barOptions.IsStacked
                            ? 1
                            : series.Count,
                        barOptions,
                        out barH,
                        out barStep,
                        out barSpan);

                    for (int ci = 0;
                         ci < categoryCount;
                         ci++)
                    {
                        for (int si = 0;
                             si < series.Count;
                             si++)
                        {
                            if (ci >=
                                series[si].Values.Count)
                            {
                                continue;
                            }

                            double value =
                                series[si].Values[ci];

                            double segmentStart = 0.0;
                            double segmentEnd = value;

                            if (barOptions.IsStacked)
                            {
                                GetStackedChartSegment(
                                    series,
                                    ci,
                                    si,
                                    barOptions.IsPercentStacked,
                                    out segmentStart,
                                    out segmentEnd);
                            }

                            float startX =
                                plot.Left +
                                (float)(
                                    ChartAxisFraction(
                                        segmentStart,
                                        axisScale) *
                                    plot.Width);

                            float valueX =
                                plot.Left +
                                (float)(
                                    ChartAxisFraction(
                                        segmentEnd,
                                        axisScale) *
                                    plot.Width);

                            float left =
                                Math.Min(
                                    startX,
                                    valueX);

                            float width =
                                Math.Max(
                                    1f,
                                    Math.Abs(
                                        valueX -
                                        startX));

                            float y =
                                plot.Top +
                                ci * groupH +
                                (groupH -
                                 barSpan) /
                                2f +
                                (barOptions.IsStacked
                                    ? 0f
                                    : si * barStep);

                            Color color =
                                GetChartSeriesColor(
                                    series,
                                    si,
                                    palette);

                            using (Brush brush =
                                new SolidBrush(color))
                            {
                                g.FillRectangle(
                                    brush,
                                    left,
                                    y,
                                    width,
                                    Math.Max(
                                        1f,
                                        barH - 1));
                            }

                            ChartLabelOptions pointLabels =
                                ResolveChartPointLabelOptions(
                                    labelOptions,
                                    series[si],
                                    ci);

                            if (pointLabels.HasAny)
                            {
                                string label =
                                    BuildChartDataLabel(
                                        pointLabels,
                                        series[si],
                                        ci,
                                        false);

                                DrawBarChartDataLabel(
                                    g,
                                    valueFont,
                                    valueBrush,
                                    label,
                                    new RectangleF(
                                        left,
                                        y,
                                        width,
                                        Math.Max(
                                            1f,
                                            barH - 1f)),
                                    startX,
                                    valueX,
                                    value,
                                    pointLabels);
                            }
                        }
                    }
                }
                else
                {
                    float groupW =
                        plot.Width /
                        categoryCount;

                    float barW;
                    float barStep;
                    float barSpan;

                    CalculateChartBarLayout(
                        groupW,
                        barOptions.IsStacked
                            ? 1
                            : series.Count,
                        barOptions,
                        out barW,
                        out barStep,
                        out barSpan);

                    for (int ci = 0;
                         ci < categoryCount;
                         ci++)
                    {
                        for (int si = 0;
                             si < series.Count;
                             si++)
                        {
                            if (ci >=
                                series[si].Values.Count)
                            {
                                continue;
                            }

                            double value =
                                series[si].Values[ci];

                            double segmentStart = 0.0;
                            double segmentEnd = value;

                            if (barOptions.IsStacked)
                            {
                                GetStackedChartSegment(
                                    series,
                                    ci,
                                    si,
                                    barOptions.IsPercentStacked,
                                    out segmentStart,
                                    out segmentEnd);
                            }

                            float startY =
                                plot.Bottom -
                                (float)(
                                    ChartAxisFraction(
                                        segmentStart,
                                        axisScale) *
                                    plot.Height);

                            float valueY =
                                plot.Bottom -
                                (float)(
                                    ChartAxisFraction(
                                        segmentEnd,
                                        axisScale) *
                                    plot.Height);

                            float top =
                                Math.Min(
                                    startY,
                                    valueY);

                            float height =
                                Math.Max(
                                    1f,
                                    Math.Abs(
                                        valueY -
                                        startY));

                            float x =
                                plot.Left +
                                ci * groupW +
                                (groupW -
                                 barSpan) /
                                2f +
                                (barOptions.IsStacked
                                    ? 0f
                                    : si * barStep);

                            Color color =
                                GetChartSeriesColor(
                                    series,
                                    si,
                                    palette);

                            using (Brush brush =
                                new SolidBrush(color))
                            {
                                g.FillRectangle(
                                    brush,
                                    x,
                                    top,
                                    Math.Max(
                                        1f,
                                        barW - 1),
                                    height);
                            }

                            ChartLabelOptions pointLabels =
                                ResolveChartPointLabelOptions(
                                    labelOptions,
                                    series[si],
                                    ci);

                            if (pointLabels.HasAny)
                            {
                                string label =
                                    BuildChartDataLabel(
                                        pointLabels,
                                        series[si],
                                        ci,
                                        false);

                                DrawColumnChartDataLabel(
                                    g,
                                    valueFont,
                                    valueBrush,
                                    label,
                                    new RectangleF(
                                        x,
                                        top,
                                        Math.Max(
                                            1f,
                                            barW - 1f),
                                        height),
                                    startY,
                                    valueY,
                                    value,
                                    pointLabels.Position);
                            }
                        }
                    }
                }
            }

            string categoryTickLabelPosition =
                ReadChartAxisTickLabelPosition(
                    chartDoc,
                    "catAx");

            if (!string.Equals(
                    categoryTickLabelPosition,
                    "none",
                    StringComparison.OrdinalIgnoreCase))
            {
                DrawChartCategoryLabels(
                    g,
                    plot,
                    series[0],
                    categoryCount,
                    kind,
                    categoryTickLabelPosition,
                    ReadChartCategoryLabelSkip(
                        chartDoc));
            }

            DrawChartAxisTitles(
                g,
                plot,
                rect,
                kind,
                valueAxisTitle,
                categoryAxisTitle);

            DrawChartLegend(
                g,
                rect,
                series,
                palette,
                kind,
                legendPosition);
        }

        private static ChartBarOptions ReadChartBarOptions(
            XmlDocument chartDoc)
        {
            ChartBarOptions result =
                new ChartBarOptions();

            XmlNode barChart =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "barChart");

            if (barChart == null)
                return result;

            XmlNode grouping =
                DirectChild(
                    barChart,
                    "grouping");

            string groupingValue =
                grouping != null
                    ? GetAttr(
                        grouping,
                        "val")
                    : null;

            if (!string.IsNullOrEmpty(
                    groupingValue))
            {
                result.Grouping =
                    groupingValue;
            }

            float parsed;

            XmlNode gapWidth =
                DirectChild(
                    barChart,
                    "gapWidth");

            if (gapWidth != null &&
                float.TryParse(
                    GetAttr(
                        gapWidth,
                        "val"),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out parsed))
            {
                result.GapWidth =
                    Math.Max(
                        0f,
                        Math.Min(
                            500f,
                            parsed));
            }

            XmlNode overlap =
                DirectChild(
                    barChart,
                    "overlap");

            if (overlap != null &&
                float.TryParse(
                    GetAttr(
                        overlap,
                        "val"),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out parsed))
            {
                result.Overlap =
                    Math.Max(
                        -100f,
                        Math.Min(
                            100f,
                            parsed));
            }

            return result;
        }

        private static void CalculateChartBarLayout(
            float slotSize,
            int seriesCount,
            ChartBarOptions options,
            out float barSize,
            out float seriesStep,
            out float groupSpan)
        {
            slotSize =
                Math.Max(
                    1f,
                    slotSize);

            seriesCount =
                Math.Max(
                    1,
                    seriesCount);

            float gapWidth =
                options != null
                    ? options.GapWidth
                    : 150f;

            float overlap =
                options != null
                    ? options.Overlap
                    : 0f;

            float overlapFactor =
                Math.Max(
                    0f,
                    Math.Min(
                        2f,
                        1f -
                        overlap /
                        100f));

            float denominator =
                1f +
                (seriesCount -
                 1) *
                overlapFactor +
                gapWidth /
                100f;

            barSize =
                Math.Max(
                    1f,
                    Math.Min(
                        slotSize,
                        slotSize /
                        Math.Max(
                            0.25f,
                            denominator)));

            seriesStep =
                barSize *
                overlapFactor;

            groupSpan =
                barSize +
                Math.Max(
                    0,
                    seriesCount -
                    1) *
                seriesStep;

            if (groupSpan >
                slotSize *
                0.98f)
            {
                float scale =
                    slotSize *
                    0.98f /
                    Math.Max(
                        1f,
                        groupSpan);

                barSize *= scale;
                seriesStep *= scale;
                groupSpan *= scale;
            }
        }

        private static void CalculateStackedChartRange(
            List<ChartSeriesData> series,
            int categoryCount,
            bool percent,
            out double minimum,
            out double maximum)
        {
            minimum = 0.0;
            maximum = 0.0;

            for (int ci = 0;
                 ci < categoryCount;
                 ci++)
            {
                double positive = 0.0;
                double negative = 0.0;

                for (int si = 0;
                     si < series.Count;
                     si++)
                {
                    if (ci >=
                        series[si].Values.Count)
                    {
                        continue;
                    }

                    double value =
                        series[si].Values[ci];

                    if (value >= 0.0)
                        positive += value;
                    else
                        negative += value;
                }

                if (percent)
                {
                    if (positive > 0.0)
                        maximum = 1.0;

                    if (negative < 0.0)
                        minimum = -1.0;
                }
                else
                {
                    maximum =
                        Math.Max(
                            maximum,
                            positive);
                    minimum =
                        Math.Min(
                            minimum,
                            negative);
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

        private static void GetStackedChartSegment(
            List<ChartSeriesData> series,
            int categoryIndex,
            int seriesIndex,
            bool percent,
            out double start,
            out double end)
        {
            start = 0.0;
            end = 0.0;

            if (series == null ||
                seriesIndex < 0 ||
                seriesIndex >= series.Count ||
                categoryIndex < 0 ||
                categoryIndex >=
                    series[seriesIndex].Values.Count)
            {
                return;
            }

            double raw =
                series[seriesIndex]
                    .Values[categoryIndex];

            double positiveTotal = 0.0;
            double negativeTotal = 0.0;

            if (percent)
            {
                for (int si = 0;
                     si < series.Count;
                     si++)
                {
                    if (categoryIndex >=
                        series[si].Values.Count)
                    {
                        continue;
                    }

                    double value =
                        series[si]
                            .Values[categoryIndex];

                    if (value >= 0.0)
                        positiveTotal += value;
                    else
                        negativeTotal +=
                            Math.Abs(value);
                }
            }

            for (int si = 0;
                 si < seriesIndex;
                 si++)
            {
                if (categoryIndex >=
                    series[si].Values.Count)
                {
                    continue;
                }

                double previous =
                    series[si]
                        .Values[categoryIndex];

                if ((raw >= 0.0 &&
                     previous < 0.0) ||
                    (raw < 0.0 &&
                     previous >= 0.0))
                {
                    continue;
                }

                if (percent)
                {
                    double total =
                        previous >= 0.0
                            ? positiveTotal
                            : negativeTotal;

                    if (total >
                        0.0000001)
                    {
                        start +=
                            previous /
                            total;
                    }
                }
                else
                {
                    start += previous;
                }
            }

            double displayValue =
                raw;

            if (percent)
            {
                double total =
                    raw >= 0.0
                        ? positiveTotal
                        : negativeTotal;

                displayValue =
                    total >
                    0.0000001
                        ? raw /
                          total
                        : 0.0;
            }

            end =
                start +
                displayValue;
        }

        private static float ReadChartFirstSliceAngle(
            XmlDocument chartDoc)
        {
            XmlNode angle =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "firstSliceAng");

            float value;

            if (angle != null &&
                float.TryParse(
                    GetAttr(
                        angle,
                        "val"),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out value))
            {
                value =
                    Math.Max(
                        0f,
                        Math.Min(
                            360f,
                            value));

                return -90f +
                    value;
            }

            return -90f;
        }

        private static float ReadDoughnutHoleRatio(
            XmlDocument chartDoc)
        {
            XmlNode hole =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "holeSize");

            float value;

            if (hole != null &&
                float.TryParse(
                    GetAttr(
                        hole,
                        "val"),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out value))
            {
                value =
                    Math.Max(
                        10f,
                        Math.Min(
                            90f,
                            value));

                return value /
                    100f;
            }

            return 0.50f;
        }

        private static XmlNode FindChartLevelDataLabels(
            XmlDocument chartDoc)
        {
            if (chartDoc == null)
                return null;

            string[] chartNames =
                new string[]
                {
                    "barChart",
                    "lineChart",
                    "pieChart",
                    "doughnutChart",
                    "areaChart",
                    "scatterChart",
                    "bubbleChart",
                    "radarChart"
                };

            for (int i = 0;
                 i < chartNames.Length;
                 i++)
            {
                XmlNode chart =
                    FindFirst(
                        chartDoc,
                        chartNames[i]);

                if (chart == null)
                    continue;

                XmlNode labels =
                    DirectChild(
                        chart,
                        "dLbls");

                if (labels != null)
                    return labels;
            }

            return null;
        }

        private static ChartLabelOptions ReadChartLabelOptions(
            XmlDocument chartDoc)
        {
            return ReadChartLabelOptions(
                chartDoc,
                null);
        }

        private static ChartLabelOptions ReadChartLabelOptions(
            XmlDocument chartDoc,
            Dictionary<string, Color> theme)
        {
            ChartLabelOptions result =
                new ChartLabelOptions();

            XmlNode labels =
                FindChartLevelDataLabels(
                    chartDoc);

            if (labels == null)
                return result;

            result.ShowValue =
                ReadChartBooleanChild(
                    labels,
                    "showVal");
            result.ShowCategoryName =
                ReadChartBooleanChild(
                    labels,
                    "showCatName");
            result.ShowSeriesName =
                ReadChartBooleanChild(
                    labels,
                    "showSerName");
            result.ShowPercent =
                ReadChartBooleanChild(
                    labels,
                    "showPercent");
            result.ShowLeaderLines =
                ReadChartBooleanChild(
                    labels,
                    "showLeaderLines");
            result.TextColor =
                ReadChartDataLabelTextColor(
                    labels,
                    theme);

            XmlNode numberFormat =
                DirectChild(
                    labels,
                    "numFmt");

            if (numberFormat != null)
            {
                result.NumberFormat =
                    GetAttr(
                        numberFormat,
                        "formatCode");
            }

            XmlNode position =
                DirectChild(
                    labels,
                    "dLblPos");

            if (position != null)
            {
                result.Position =
                    GetAttr(
                        position,
                        "val");
            }

            XmlNode separator =
                DirectChild(
                    labels,
                    "separator");

            if (separator != null &&
                !string.IsNullOrEmpty(
                    separator.InnerText))
            {
                result.Separator =
                    separator.InnerText;
            }

            return result;
        }

        private static void ReadChartSeriesLabelOverrides(
            XmlNode series,
            ChartSeriesData data)
        {
            ReadChartSeriesLabelOverrides(
                series,
                data,
                null);
        }

        private static void ReadChartSeriesLabelOverrides(
            XmlNode series,
            ChartSeriesData data,
            Dictionary<string, Color> theme)
        {
            if (series == null ||
                data == null)
            {
                return;
            }

            XmlNode labels =
                DirectChild(
                    series,
                    "dLbls");

            if (labels == null)
                return;

            for (int i = 0;
                 i < labels.ChildNodes.Count;
                 i++)
            {
                XmlNode label =
                    labels.ChildNodes[i];

                if (label == null ||
                    label.NodeType !=
                        XmlNodeType.Element ||
                    label.LocalName !=
                        "dLbl")
                {
                    continue;
                }

                XmlNode index =
                    DirectChild(
                        label,
                        "idx");

                int pointIndex;

                if (index == null ||
                    !int.TryParse(
                        GetAttr(
                            index,
                            "val"),
                        out pointIndex) ||
                    pointIndex < 0)
                {
                    continue;
                }

                ChartPointLabelOverride item =
                    new ChartPointLabelOverride();

                bool parsedBoolean;

                if (TryReadChartBooleanChild(
                        label,
                        "showVal",
                        out parsedBoolean))
                {
                    item.ShowValue =
                        parsedBoolean;
                }

                if (TryReadChartBooleanChild(
                        label,
                        "showCatName",
                        out parsedBoolean))
                {
                    item.ShowCategoryName =
                        parsedBoolean;
                }

                if (TryReadChartBooleanChild(
                        label,
                        "showSerName",
                        out parsedBoolean))
                {
                    item.ShowSeriesName =
                        parsedBoolean;
                }

                if (TryReadChartBooleanChild(
                        label,
                        "showPercent",
                        out parsedBoolean))
                {
                    item.ShowPercent =
                        parsedBoolean;
                }

                if (TryReadChartBooleanChild(
                        label,
                        "delete",
                        out parsedBoolean))
                {
                    item.Delete =
                        parsedBoolean;
                }

                item.TextColor =
                    ReadChartDataLabelTextColor(
                        label,
                        theme);

                XmlNode numberFormat =
                    DirectChild(
                        label,
                        "numFmt");

                if (numberFormat != null)
                {
                    item.NumberFormat =
                        GetAttr(
                            numberFormat,
                            "formatCode");
                }

                XmlNode position =
                    DirectChild(
                        label,
                        "dLblPos");

                if (position != null)
                {
                    item.Position =
                        GetAttr(
                            position,
                            "val");
                }

                XmlNode separator =
                    DirectChild(
                        label,
                        "separator");

                if (separator != null)
                {
                    item.Separator =
                        separator.InnerText;
                }

                data.PointLabelOverrides[
                    pointIndex] =
                    item;
            }
        }

        private static Color? ReadChartDataLabelTextColor(
            XmlNode labels,
            Dictionary<string, Color> theme)
        {
            if (labels == null)
                return null;

            XmlNode textProperties =
                DirectChild(
                    labels,
                    "txPr");

            if (textProperties == null)
                return null;

            XmlNode runProperties =
                FindFirst(
                    textProperties,
                    "defRPr");

            if (runProperties == null)
            {
                runProperties =
                    FindFirst(
                        textProperties,
                        "rPr");
            }

            if (runProperties == null)
            {
                runProperties =
                    FindFirst(
                        textProperties,
                        "endParaRPr");
            }

            return ReadSolidFill(
                runProperties ??
                textProperties,
                theme);
        }

        private static bool TryReadChartBooleanChild(
            XmlNode parent,
            string childName,
            out bool value)
        {
            value = false;

            XmlNode child =
                DirectChild(
                    parent,
                    childName);

            if (child == null)
                return false;

            string raw =
                GetAttr(
                    child,
                    "val");

            value =
                string.IsNullOrEmpty(
                    raw) ||
                raw == "1" ||
                string.Equals(
                    raw,
                    "true",
                    StringComparison.OrdinalIgnoreCase);

            return true;
        }

        private static ChartLabelOptions ResolveChartPointLabelOptions(
            ChartLabelOptions defaults,
            ChartSeriesData series,
            int index)
        {
            ChartLabelOptions result =
                new ChartLabelOptions();

            if (defaults != null)
            {
                result.ShowValue =
                    defaults.ShowValue;
                result.ShowCategoryName =
                    defaults.ShowCategoryName;
                result.ShowSeriesName =
                    defaults.ShowSeriesName;
                result.ShowPercent =
                    defaults.ShowPercent;
                result.ShowLeaderLines =
                    defaults.ShowLeaderLines;
                result.TextColor =
                    defaults.TextColor;
                result.NumberFormat =
                    defaults.NumberFormat;
                result.Position =
                    defaults.Position;
                result.Separator =
                    defaults.Separator;
            }

            if (series == null)
                return result;

            ChartPointLabelOverride item;

            if (!series.PointLabelOverrides.TryGetValue(
                    index,
                    out item) ||
                item == null)
            {
                return result;
            }

            if (item.Delete)
            {
                result.ShowValue = false;
                result.ShowCategoryName = false;
                result.ShowSeriesName = false;
                result.ShowPercent = false;
                return result;
            }

            if (item.ShowValue.HasValue)
                result.ShowValue =
                    item.ShowValue.Value;

            if (item.ShowCategoryName.HasValue)
                result.ShowCategoryName =
                    item.ShowCategoryName.Value;

            if (item.ShowSeriesName.HasValue)
                result.ShowSeriesName =
                    item.ShowSeriesName.Value;

            if (item.ShowPercent.HasValue)
                result.ShowPercent =
                    item.ShowPercent.Value;

            if (item.TextColor.HasValue)
                result.TextColor =
                    item.TextColor.Value;

            if (!string.IsNullOrEmpty(
                    item.NumberFormat))
            {
                result.NumberFormat =
                    item.NumberFormat;
            }

            if (!string.IsNullOrEmpty(
                    item.Position))
            {
                result.Position =
                    item.Position;
            }

            if (item.Separator != null)
            {
                result.Separator =
                    item.Separator;
            }

            return result;
        }

        private static bool HasAnyChartSeriesDataLabel(
            ChartLabelOptions defaults,
            ChartSeriesData series)
        {
            if (defaults != null &&
                defaults.HasAny)
            {
                return true;
            }

            if (series == null)
                return false;

            foreach (KeyValuePair<int, ChartPointLabelOverride> pair
                in series.PointLabelOverrides)
            {
                ChartLabelOptions resolved =
                    ResolveChartPointLabelOptions(
                        defaults,
                        series,
                        pair.Key);

                if (resolved.HasAny)
                    return true;
            }

            return false;
        }

        private static bool ReadChartBooleanChild(
            XmlNode parent,
            string childName)
        {
            XmlNode child =
                DirectChild(
                    parent,
                    childName);

            if (child == null)
                return false;

            string value =
                GetAttr(
                    child,
                    "val");

            if (string.IsNullOrEmpty(value))
                return true;

            return value == "1" ||
                string.Equals(
                    value,
                    "true",
                    StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildChartDataLabel(
            ChartLabelOptions options,
            ChartSeriesData series,
            int index,
            bool pie)
        {
            if (options == null ||
                series == null ||
                index < 0 ||
                index >= series.Values.Count)
            {
                return string.Empty;
            }

            List<string> parts =
                new List<string>();

            if (options.ShowSeriesName &&
                !string.IsNullOrEmpty(
                    series.Name))
            {
                parts.Add(
                    series.Name);
            }

            if (options.ShowCategoryName &&
                index <
                series.Categories.Count &&
                !string.IsNullOrEmpty(
                    series.Categories[index]))
            {
                parts.Add(
                    series.Categories[index]);
            }

            if (options.ShowValue)
            {
                parts.Add(
                    FormatChartAxisNumber(
                        series.Values[index],
                        options.NumberFormat));
            }

            if (options.ShowPercent &&
                pie)
            {
                double total = 0.0;

                for (int i = 0;
                     i < series.Values.Count;
                     i++)
                {
                    total +=
                        Math.Abs(
                            series.Values[i]);
                }

                if (total > 0.0000001)
                {
                    double ratio =
                        Math.Abs(
                            series.Values[index]) /
                        total;

                    parts.Add(
                        ratio.ToString(
                            "0.#%",
                            System.Globalization.CultureInfo.InvariantCulture));
                }
            }

            if (parts.Count == 0)
                return string.Empty;

            return string.Join(
                string.IsNullOrEmpty(
                    options.Separator)
                    ? ", "
                    : options.Separator,
                parts.ToArray());
        }

        private static string ReadChartLegendPosition(
            XmlDocument chartDoc)
        {
            XmlNode legend =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "legend");

            if (legend == null)
                return string.Empty;

            XmlNode deleted =
                DirectChild(
                    legend,
                    "delete");

            if (deleted != null &&
                ReadChartBooleanChild(
                    legend,
                    "delete"))
            {
                return string.Empty;
            }

            XmlNode position =
                DirectChild(
                    legend,
                    "legendPos");

            string value =
                position != null
                    ? GetAttr(
                        position,
                        "val")
                    : string.Empty;

            if (value == "l" ||
                value == "r" ||
                value == "t" ||
                value == "b" ||
                value == "tr")
            {
                return value;
            }

            return "r";
        }

        private static ChartAxisScale ReadChartAxisScale(
            XmlDocument chartDoc,
            double dataMinimum,
            double dataMaximum)
        {
            ChartAxisScale scale =
                new ChartAxisScale();

            scale.Minimum =
                dataMinimum;
            scale.Maximum =
                dataMaximum;
            scale.MajorUnit = 0.0;
            scale.Reverse = false;

            XmlNode valueAxis =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "valAx");

            if (valueAxis == null)
            {
                NormalizeChartAxisScale(scale);
                return scale;
            }

            XmlNode scaling =
                DirectChild(
                    valueAxis,
                    "scaling");

            if (scaling != null)
            {
                XmlNode minimum =
                    DirectChild(
                        scaling,
                        "min");
                XmlNode maximum =
                    DirectChild(
                        scaling,
                        "max");
                XmlNode orientation =
                    DirectChild(
                        scaling,
                        "orientation");

                double parsed;

                if (minimum != null &&
                    double.TryParse(
                        GetAttr(
                            minimum,
                            "val"),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out parsed))
                {
                    scale.Minimum =
                        parsed;
                }

                if (maximum != null &&
                    double.TryParse(
                        GetAttr(
                            maximum,
                            "val"),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out parsed))
                {
                    scale.Maximum =
                        parsed;
                }

                if (orientation != null)
                {
                    scale.Reverse =
                        string.Equals(
                            GetAttr(
                                orientation,
                                "val"),
                            "maxMin",
                            StringComparison.OrdinalIgnoreCase);
                }
            }

            XmlNode majorUnit =
                DirectChild(
                    valueAxis,
                    "majorUnit");

            double unit;
            if (majorUnit != null &&
                double.TryParse(
                    GetAttr(
                        majorUnit,
                        "val"),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out unit) &&
                unit > 0.0 &&
                !double.IsNaN(unit) &&
                !double.IsInfinity(unit))
            {
                scale.MajorUnit =
                    unit;
            }

            XmlNode minorUnit =
                DirectChild(
                    valueAxis,
                    "minorUnit");

            if (minorUnit != null &&
                double.TryParse(
                    GetAttr(
                        minorUnit,
                        "val"),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out unit) &&
                unit > 0.0 &&
                !double.IsNaN(unit) &&
                !double.IsInfinity(unit))
            {
                scale.MinorUnit =
                    unit;
            }

            XmlNode numberFormat =
                DirectChild(
                    valueAxis,
                    "numFmt");

            if (numberFormat != null)
            {
                scale.NumberFormat =
                    GetAttr(
                        numberFormat,
                        "formatCode");
            }

            NormalizeChartAxisScale(scale);
            return scale;
        }

        private static void NormalizeChartAxisScale(
            ChartAxisScale scale)
        {
            if (scale == null)
                return;

            if (double.IsNaN(scale.Minimum) ||
                double.IsInfinity(scale.Minimum))
            {
                scale.Minimum = 0.0;
            }

            if (double.IsNaN(scale.Maximum) ||
                double.IsInfinity(scale.Maximum))
            {
                scale.Maximum =
                    scale.Minimum + 1.0;
            }

            if (scale.Maximum <= scale.Minimum)
            {
                scale.Maximum =
                    scale.Minimum + 1.0;
            }

            double range =
                scale.Maximum -
                scale.Minimum;

            if (scale.MajorUnit <= 0.0 ||
                scale.MajorUnit > range * 4.0)
            {
                scale.MajorUnit = 0.0;
            }

            if (scale.MinorUnit <= 0.0 ||
                scale.MinorUnit > range * 4.0)
            {
                scale.MinorUnit = 0.0;
            }
        }

        private static double ChartAxisFraction(
            double value,
            ChartAxisScale scale)
        {
            if (scale == null)
                return 0.0;

            double range =
                Math.Max(
                    0.0000001,
                    scale.Maximum -
                    scale.Minimum);

            double fraction =
                (value -
                 scale.Minimum) /
                range;

            fraction =
                Math.Max(
                    0.0,
                    Math.Min(
                        1.0,
                        fraction));

            return scale.Reverse
                ? 1.0 - fraction
                : fraction;
        }

        private static void DrawChartValueGrid(
            Graphics g,
            RectangleF plot,
            ChartAxisScale scale,
            string kind,
            bool showLabels,
            string tickLabelPosition,
            ChartLineStyle gridStyle,
            ChartLineStyle minorGridStyle)
        {
            if (scale == null)
                return;

            List<double> ticks =
                BuildChartAxisTicks(
                    scale);

            if (minorGridStyle != null)
            {
                List<double> minorTicks =
                    BuildChartMinorAxisTicks(
                        scale,
                        ticks);

                using (Pen minorGrid =
                    new Pen(
                        minorGridStyle.Color,
                        minorGridStyle.Width))
                {
                    minorGrid.DashStyle =
                        minorGridStyle.DashStyle;

                    for (int i = 0;
                         i < minorTicks.Count;
                         i++)
                    {
                        double fraction =
                            ChartAxisFraction(
                                minorTicks[i],
                                scale);

                        if (kind == "bar")
                        {
                            float x =
                                plot.Left +
                                plot.Width *
                                (float)fraction;

                            g.DrawLine(
                                minorGrid,
                                x,
                                plot.Top,
                                x,
                                plot.Bottom);
                        }
                        else
                        {
                            float y =
                                plot.Bottom -
                                plot.Height *
                                (float)fraction;

                            g.DrawLine(
                                minorGrid,
                                plot.Left,
                                y,
                                plot.Right,
                                y);
                        }
                    }
                }
            }

            if (gridStyle == null)
            {
                gridStyle =
                    new ChartLineStyle();
                gridStyle.Color =
                    Color.FromArgb(
                        225,
                        228,
                        232);
                gridStyle.Width = 1f;
            }

            using (Pen grid =
                new Pen(
                    gridStyle.Color,
                    gridStyle.Width))
            using (Font font = SafeFont(
                "Arial",
                Math.Max(
                    6f,
                    Math.Min(
                        9f,
                        plot.Height / 38f))))
            using (Brush text =
                new SolidBrush(
                    Color.FromArgb(
                        105,
                        105,
                        105)))
            {
                grid.DashStyle =
                    gridStyle.DashStyle;

                bool high =
                    string.Equals(
                        tickLabelPosition,
                        "high",
                        StringComparison.OrdinalIgnoreCase);

                for (int i = 0;
                     i < ticks.Count;
                     i++)
                {
                    double value =
                        ticks[i];

                    double fraction =
                        ChartAxisFraction(
                            value,
                            scale);

                    if (kind == "bar")
                    {
                        float x =
                            plot.Left +
                            plot.Width *
                            (float)fraction;

                        g.DrawLine(
                            grid,
                            x,
                            plot.Top,
                            x,
                            plot.Bottom);

                        if (showLabels)
                        {
                            string label =
                                FormatChartAxisNumber(
                                    value,
                                    scale.NumberFormat);

                            SizeF size =
                                g.MeasureString(
                                    label,
                                    font);

                            g.DrawString(
                                label,
                                font,
                                text,
                                x -
                                    size.Width /
                                    2f,
                                high
                                    ? plot.Top -
                                      size.Height -
                                      2f
                                    : plot.Bottom +
                                      2f);
                        }
                    }
                    else
                    {
                        float y =
                            plot.Bottom -
                            plot.Height *
                            (float)fraction;

                        g.DrawLine(
                            grid,
                            plot.Left,
                            y,
                            plot.Right,
                            y);

                        if (showLabels)
                        {
                            string label =
                                FormatChartAxisNumber(
                                    value,
                                    scale.NumberFormat);

                            SizeF size =
                                g.MeasureString(
                                    label,
                                    font);

                            g.DrawString(
                                label,
                                font,
                                text,
                                high
                                    ? plot.Right +
                                      4f
                                    : plot.Left -
                                      size.Width -
                                      4f,
                                y -
                                    size.Height /
                                    2f);
                        }
                    }
                }
            }
        }

        private static List<double> BuildChartAxisTicks(
            ChartAxisScale scale)
        {
            List<double> ticks =
                new List<double>();

            if (scale == null)
                return ticks;

            double range =
                scale.Maximum -
                scale.Minimum;

            double unit =
                scale.MajorUnit;

            if (unit <= 0.0)
            {
                unit =
                    ChooseNiceChartUnit(
                        range /
                        5.0);
            }

            if (unit <= 0.0)
                unit = 1.0;

            double first =
                Math.Ceiling(
                    scale.Minimum /
                    unit) *
                unit;

            int guard = 0;

            for (double value = first;
                 value <=
                    scale.Maximum +
                    unit * 0.0001 &&
                 guard < 64;
                 value += unit)
            {
                ticks.Add(value);
                guard++;
            }

            if (ticks.Count == 0)
            {
                ticks.Add(
                    scale.Minimum);
                ticks.Add(
                    scale.Maximum);
            }

            return ticks;
        }

        private static List<double> BuildChartMinorAxisTicks(
            ChartAxisScale scale,
            List<double> majorTicks)
        {
            List<double> ticks =
                new List<double>();

            if (scale == null)
                return ticks;

            double unit =
                scale.MinorUnit;

            if (unit <= 0.0)
            {
                double major =
                    scale.MajorUnit;

                if (major <= 0.0)
                {
                    major =
                        ChooseNiceChartUnit(
                            (scale.Maximum -
                             scale.Minimum) /
                            5.0);
                }

                unit =
                    major /
                    2.0;
            }

            if (unit <= 0.0)
                return ticks;

            double first =
                Math.Ceiling(
                    scale.Minimum /
                    unit) *
                unit;

            int guard = 0;

            for (double value = first;
                 value <=
                    scale.Maximum +
                    unit * 0.0001 &&
                 guard < 128;
                 value += unit)
            {
                bool isMajor =
                    false;

                if (majorTicks != null)
                {
                    for (int i = 0;
                         i < majorTicks.Count;
                         i++)
                    {
                        if (Math.Abs(
                                majorTicks[i] -
                                value) <=
                            Math.Max(
                                0.0000001,
                                unit *
                                0.001))
                        {
                            isMajor = true;
                            break;
                        }
                    }
                }

                if (!isMajor)
                {
                    ticks.Add(
                        value);
                }

                guard++;
            }

            return ticks;
        }

        private static double ChooseNiceChartUnit(
            double raw)
        {
            if (raw <= 0.0 ||
                double.IsNaN(raw) ||
                double.IsInfinity(raw))
            {
                return 1.0;
            }

            double exponent =
                Math.Floor(
                    Math.Log10(raw));

            double magnitude =
                Math.Pow(
                    10.0,
                    exponent);

            double normalized =
                raw /
                magnitude;

            double nice;

            if (normalized <= 1.0)
                nice = 1.0;
            else if (normalized <= 2.0)
                nice = 2.0;
            else if (normalized <= 5.0)
                nice = 5.0;
            else
                nice = 10.0;

            return nice *
                magnitude;
        }

        private static void DrawChartLabelText(
            Graphics g,
            string text,
            Font font,
            Brush fallbackBrush,
            float x,
            float y,
            ChartLabelOptions options)
        {
            SolidBrush ownedBrush =
                null;

            try
            {
                Brush drawBrush =
                    fallbackBrush;

                if (options != null &&
                    options.TextColor.HasValue)
                {
                    ownedBrush =
                        new SolidBrush(
                            options.TextColor.Value);
                    drawBrush =
                        ownedBrush;
                }

                g.DrawString(
                    text,
                    font,
                    drawBrush,
                    x,
                    y);
            }
            finally
            {
                if (ownedBrush != null)
                    ownedBrush.Dispose();
            }
        }

        private static string NormalizeChartDataLabelPosition(
            string position)
        {
            return string.IsNullOrEmpty(
                    position)
                ? string.Empty
                : position.Trim()
                    .ToLowerInvariant();
        }

        private static void DrawChartPointDataLabel(
            Graphics g,
            Font font,
            Brush brush,
            string text,
            PointF point,
            string position)
        {
            ChartLabelOptions options =
                new ChartLabelOptions();
            options.Position =
                position;

            DrawChartPointDataLabel(
                g,
                font,
                brush,
                text,
                point,
                options);
        }

        private static void DrawChartPointDataLabel(
            Graphics g,
            Font font,
            Brush brush,
            string text,
            PointF point,
            ChartLabelOptions options)
        {
            if (string.IsNullOrEmpty(text))
                return;

            string normalized =
                NormalizeChartDataLabelPosition(
                    options == null
                        ? null
                        : options.Position);

            SizeF size =
                g.MeasureString(
                    text,
                    font);

            float x =
                point.X -
                size.Width /
                2f;
            float y =
                point.Y -
                size.Height -
                4f;

            if (normalized == "b")
            {
                y =
                    point.Y +
                    4f;
            }
            else if (normalized == "l")
            {
                x =
                    point.X -
                    size.Width -
                    4f;
                y =
                    point.Y -
                    size.Height /
                    2f;
            }
            else if (normalized == "r")
            {
                x =
                    point.X +
                    4f;
                y =
                    point.Y -
                    size.Height /
                    2f;
            }
            else if (normalized == "ctr")
            {
                y =
                    point.Y -
                    size.Height /
                    2f;
            }

            DrawChartLabelText(
                g,
                text,
                font,
                brush,
                x,
                y,
                options);
        }

        private static void DrawBarChartDataLabel(
            Graphics g,
            Font font,
            Brush brush,
            string text,
            RectangleF bar,
            float baseX,
            float valueX,
            double value,
            string position)
        {
            ChartLabelOptions options =
                new ChartLabelOptions();
            options.Position =
                position;

            DrawBarChartDataLabel(
                g,
                font,
                brush,
                text,
                bar,
                baseX,
                valueX,
                value,
                options);
        }

        private static void DrawBarChartDataLabel(
            Graphics g,
            Font font,
            Brush brush,
            string text,
            RectangleF bar,
            float baseX,
            float valueX,
            double value,
            ChartLabelOptions options)
        {
            if (string.IsNullOrEmpty(text))
                return;

            string normalized =
                NormalizeChartDataLabelPosition(
                    options == null
                        ? null
                        : options.Position);

            SizeF size =
                g.MeasureString(
                    text,
                    font);
            float centerY =
                bar.Top +
                (bar.Height -
                 size.Height) /
                2f;

            float x;

            if (normalized == "ctr")
            {
                x =
                    bar.Left +
                    (bar.Width -
                     size.Width) /
                    2f;
            }
            else if (normalized == "inbase")
            {
                x =
                    value >= 0.0
                        ? baseX + 4f
                        : baseX -
                          size.Width -
                          4f;
            }
            else if (normalized == "inend")
            {
                x =
                    value >= 0.0
                        ? valueX -
                          size.Width -
                          4f
                        : valueX + 4f;
            }
            else
            {
                x =
                    value >= 0.0
                        ? valueX + 4f
                        : valueX -
                          size.Width -
                          4f;
            }

            DrawChartLabelText(
                g,
                text,
                font,
                brush,
                x,
                centerY,
                options);
        }

        private static void DrawColumnChartDataLabel(
            Graphics g,
            Font font,
            Brush brush,
            string text,
            RectangleF bar,
            float baseY,
            float valueY,
            double value,
            string position)
        {
            ChartLabelOptions options =
                new ChartLabelOptions();
            options.Position =
                position;

            DrawColumnChartDataLabel(
                g,
                font,
                brush,
                text,
                bar,
                baseY,
                valueY,
                value,
                options);
        }

        private static void DrawColumnChartDataLabel(
            Graphics g,
            Font font,
            Brush brush,
            string text,
            RectangleF bar,
            float baseY,
            float valueY,
            double value,
            ChartLabelOptions options)
        {
            if (string.IsNullOrEmpty(text))
                return;

            string normalized =
                NormalizeChartDataLabelPosition(
                    options == null
                        ? null
                        : options.Position);

            SizeF size =
                g.MeasureString(
                    text,
                    font);

            float x =
                bar.Left +
                (bar.Width -
                 size.Width) /
                2f;
            float y;

            if (normalized == "ctr")
            {
                y =
                    bar.Top +
                    (bar.Height -
                     size.Height) /
                    2f;
            }
            else if (normalized == "inbase")
            {
                y =
                    value >= 0.0
                        ? baseY -
                          size.Height -
                          2f
                        : baseY + 2f;
            }
            else if (normalized == "inend")
            {
                y =
                    value >= 0.0
                        ? valueY + 2f
                        : valueY -
                          size.Height -
                          2f;
            }
            else if (normalized == "b")
            {
                y =
                    bar.Bottom +
                    2f;
            }
            else if (normalized == "t")
            {
                y =
                    bar.Top -
                    size.Height -
                    2f;
            }
            else
            {
                y =
                    value >= 0.0
                        ? valueY -
                          size.Height -
                          2f
                        : valueY + 2f;
            }

            DrawChartLabelText(
                g,
                text,
                font,
                brush,
                x,
                y,
                options);
        }

        private static void DrawChartValueLabel(
            Graphics g,
            Font font,
            Brush brush,
            string text,
            float x,
            float y,
            bool centered)
        {
            if (string.IsNullOrEmpty(text))
                return;

            SizeF size =
                g.MeasureString(
                    text,
                    font);

            float drawX =
                centered
                    ? x - size.Width / 2f
                    : x;

            g.DrawString(
                text,
                font,
                brush,
                drawX,
                y);
        }

        private static string FormatChartNumber(
            double value)
        {
            double abs =
                Math.Abs(value);

            if (abs >= 1000000.0)
            {
                return (
                    value /
                    1000000.0)
                    .ToString(
                        "0.##",
                        System.Globalization.CultureInfo.InvariantCulture) +
                    "M";
            }

            if (abs >= 1000.0)
            {
                return (
                    value /
                    1000.0)
                    .ToString(
                        "0.##",
                        System.Globalization.CultureInfo.InvariantCulture) +
                    "K";
            }

            return value.ToString(
                "0.##",
                System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void DrawPieChartValueLabels(
            Graphics g,
            RectangleF plot,
            ChartSeriesData series,
            ChartLabelOptions options,
            float startAngle,
            ChartLineStyle leaderLineStyle)
        {
            if (series == null ||
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

            if (total <= 0.0)
                return;

            float diameter =
                Math.Min(
                    plot.Width,
                    plot.Height) *
                0.82f;

            float cx =
                plot.Left +
                plot.Width /
                2f;

            float cy =
                plot.Top +
                plot.Height /
                2f;

            float start =
                startAngle;

            using (Font font = SafeFont(
                "Arial",
                Math.Max(
                    7f,
                    Math.Min(
                        10f,
                        plot.Height /
                        30f)),
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

                    float angle =
                        start +
                        sweep /
                        2f;

                    double radians =
                        angle *
                        Math.PI /
                        180.0;

                    ChartLabelOptions pointLabels =
                        ResolveChartPointLabelOptions(
                            options,
                            series,
                            i);

                    if (pointLabels.HasAny)
                    {
                        string labelPosition =
                            NormalizeChartDataLabelPosition(
                                pointLabels.Position);

                        float radiusFactor =
                            labelPosition == "ctr"
                                ? 0.18f
                                : labelPosition == "outend"
                                    ? 0.58f
                                    : labelPosition == "inend"
                                        ? 0.31f
                                        : 0.34f;

                        float radius =
                            diameter *
                            radiusFactor;

                        string label =
                            BuildChartDataLabel(
                                pointLabels,
                                series,
                                i,
                                true);

                        SizeF size =
                            g.MeasureString(
                                label,
                                font);

                        float x =
                            cx +
                            (float)Math.Cos(
                                radians) *
                            radius -
                            size.Width /
                            2f;

                        float y =
                            cy +
                            (float)Math.Sin(
                                radians) *
                            radius -
                            size.Height /
                            2f;

                        if (labelPosition == "outend" &&
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

                            float edgeRadius =
                                diameter *
                                0.49f;
                            float lineEndRadius =
                                diameter *
                                0.545f;

                            PointF lineStart =
                                new PointF(
                                    cx +
                                    (float)Math.Cos(
                                        radians) *
                                    edgeRadius,
                                    cy +
                                    (float)Math.Sin(
                                        radians) *
                                    edgeRadius);
                            PointF lineEnd =
                                new PointF(
                                    cx +
                                    (float)Math.Cos(
                                        radians) *
                                    lineEndRadius,
                                    cy +
                                    (float)Math.Sin(
                                        radians) *
                                    lineEndRadius);

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

                    start += sweep;
                }
            }
        }

        private static void DrawChartCategoryLabels(
            Graphics g,
            RectangleF plot,
            ChartSeriesData firstSeries,
            int categoryCount,
            string kind,
            string tickLabelPosition,
            int labelSkip)
        {
            if (firstSeries == null || firstSeries.Categories.Count == 0)
                return;

            using (Font font = SafeFont("Arial", Math.Max(6f, Math.Min(9f, plot.Height / 35f))))
            using (Brush brush = new SolidBrush(Color.FromArgb(80, 80, 80)))
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Near;
                sf.Trimming = StringTrimming.EllipsisCharacter;

                int count = Math.Min(categoryCount, firstSeries.Categories.Count);
                bool high =
                    string.Equals(
                        tickLabelPosition,
                        "high",
                        StringComparison.OrdinalIgnoreCase);

                labelSkip =
                    Math.Max(
                        1,
                        labelSkip);

                for (int i = 0;
                     i < count;
                     i += labelSkip)
                {
                    if (kind == "bar")
                    {
                        sf.Alignment =
                            high
                                ? StringAlignment.Near
                                : StringAlignment.Far;
                        float groupH = plot.Height / Math.Max(1, categoryCount);
                        float labelWidth =
                            Math.Max(
                                20f,
                                plot.Width *
                                0.08f);
                        RectangleF label = new RectangleF(
                            high
                                ? plot.Right + 4f
                                : plot.Left -
                                  Math.Max(
                                      24f,
                                      plot.Width *
                                      0.09f),
                            plot.Top + i * groupH,
                            labelWidth,
                            groupH);
                        sf.LineAlignment = StringAlignment.Center;
                        g.DrawString(firstSeries.Categories[i], font, brush, label, sf);
                    }
                    else
                    {
                        sf.Alignment = StringAlignment.Center;
                        sf.LineAlignment = StringAlignment.Near;
                        float groupW = plot.Width / Math.Max(1, categoryCount);
                        float labelHeight =
                            Math.Max(
                                14f,
                                font.Height +
                                4f);
                        RectangleF label = new RectangleF(
                            plot.Left + i * groupW,
                            high
                                ? plot.Top -
                                  labelHeight -
                                  2f
                                : plot.Bottom +
                                  2f,
                            groupW,
                            labelHeight);
                        g.DrawString(firstSeries.Categories[i], font, brush, label, sf);
                    }
                }
            }
        }

        private static Color ReadChartAreaFill(
            XmlDocument chartDoc,
            Dictionary<string, Color> theme,
            string areaName,
            Color fallback)
        {
            if (chartDoc == null ||
                string.IsNullOrEmpty(
                    areaName))
            {
                return fallback;
            }

            XmlNode area =
                string.Equals(
                    areaName,
                    "chartSpace",
                    StringComparison.Ordinal)
                    ? chartDoc.DocumentElement
                    : FindFirst(
                        chartDoc,
                        areaName);

            if (area == null)
                return fallback;

            XmlNode shapeProperties =
                DirectChild(
                    area,
                    "spPr");

            if (shapeProperties == null)
                return fallback;

            Color? fill =
                ReadSolidFill(
                    shapeProperties,
                    theme);

            return fill.HasValue
                ? fill.Value
                : fallback;
        }

        private static Color ThemeOrDefault(
            Dictionary<string, Color> theme,
            string key,
            Color fallback)
        {
            return theme != null && theme.ContainsKey(key) ? theme[key] : fallback;
        }

        private static Color? ReadChartSeriesColor(
            XmlNode series,
            Dictionary<string, Color> theme)
        {
            if (series == null)
                return null;

            XmlNode shapeProperties =
                DirectChild(
                    series,
                    "spPr");

            if (shapeProperties == null)
                return null;

            return ReadSolidFill(
                shapeProperties,
                theme);
        }

        private static void ReadChartPointColors(
            XmlNode series,
            Dictionary<string, Color> theme,
            List<Color?> output,
            int pointCount)
        {
            output.Clear();

            for (int i = 0;
                 i < pointCount;
                 i++)
            {
                output.Add(null);
            }

            if (series == null)
                return;

            foreach (XmlNode point in
                FindAll(
                    series,
                    "dPt"))
            {
                XmlNode index =
                    DirectChild(
                        point,
                        "idx");

                int pointIndex;
                if (index == null ||
                    !int.TryParse(
                        GetAttr(
                            index,
                            "val"),
                        out pointIndex) ||
                    pointIndex < 0 ||
                    pointIndex >= output.Count)
                {
                    continue;
                }

                XmlNode shapeProperties =
                    DirectChild(
                        point,
                        "spPr");

                if (shapeProperties == null)
                    continue;

                Color? color =
                    ReadSolidFill(
                        shapeProperties,
                        theme);

                if (color.HasValue)
                {
                    output[pointIndex] =
                        color.Value;
                }
            }
        }

        private static Color GetChartSeriesColor(
            List<ChartSeriesData> series,
            int index,
            Color[] palette)
        {
            if (series != null &&
                index >= 0 &&
                index < series.Count &&
                series[index] != null &&
                series[index].ExplicitColor.HasValue)
            {
                return series[index]
                    .ExplicitColor
                    .Value;
            }

            if (palette == null ||
                palette.Length == 0)
            {
                return Color.FromArgb(
                    79,
                    129,
                    189);
            }

            return palette[
                Math.Abs(index) %
                palette.Length];
        }

        private static Color GetChartPointColor(
            ChartSeriesData series,
            int index,
            Color[] palette)
        {
            if (series != null &&
                index >= 0 &&
                index <
                series.PointColors.Count &&
                series.PointColors[index].HasValue)
            {
                return series
                    .PointColors[index]
                    .Value;
            }

            if (series != null &&
                series.ExplicitColor.HasValue)
            {
                return series
                    .ExplicitColor
                    .Value;
            }

            if (palette == null ||
                palette.Length == 0)
            {
                return Color.FromArgb(
                    79,
                    129,
                    189);
            }

            return palette[
                Math.Abs(index) %
                palette.Length];
        }

        private static void ReadChartSeriesVisualStyle(
            XmlNode series,
            ChartSeriesData data)
        {
            if (series == null ||
                data == null)
            {
                return;
            }

            XmlNode shapeProperties =
                DirectChild(
                    series,
                    "spPr");

            XmlNode line =
                shapeProperties != null
                    ? DirectChild(
                        shapeProperties,
                        "ln")
                    : null;

            if (line != null)
            {
                long width =
                    GetLong(
                        line,
                        "w",
                        0);

                if (width > 0)
                {
                    data.LineWidth =
                        Math.Max(
                            1f,
                            EmuToRenderPixels(
                                width));
                }

                XmlNode dash =
                    DirectChild(
                        line,
                        "prstDash");

                string dashValue =
                    dash != null
                        ? GetAttr(
                            dash,
                            "val")
                        : string.Empty;

                data.LineDashStyle =
                    ParseChartDashStyle(
                        dashValue);
            }

            XmlNode marker =
                DirectChild(
                    series,
                    "marker");

            if (marker == null)
                return;

            XmlNode symbol =
                DirectChild(
                    marker,
                    "symbol");

            string symbolValue =
                symbol != null
                    ? GetAttr(
                        symbol,
                        "val")
                    : string.Empty;

            if (!string.IsNullOrEmpty(
                    symbolValue))
            {
                data.MarkerSymbol =
                    symbolValue;
                data.MarkerEnabled =
                    !string.Equals(
                        symbolValue,
                        "none",
                        StringComparison.OrdinalIgnoreCase);
            }

            XmlNode size =
                DirectChild(
                    marker,
                    "size");

            int sizeValue;
            if (size != null &&
                int.TryParse(
                    GetAttr(
                        size,
                        "val"),
                    out sizeValue))
            {
                data.MarkerSize =
                    Math.Max(
                        3f,
                        Math.Min(
                            24f,
                            sizeValue));
            }
        }

        private static DashStyle ParseChartDashStyle(
            string value)
        {
            if (string.IsNullOrEmpty(value) ||
                value == "solid")
            {
                return DashStyle.Solid;
            }

            string lower =
                value.ToLowerInvariant();

            if (lower.IndexOf(
                    "dashdot") >= 0)
            {
                return DashStyle.DashDot;
            }

            if (lower.IndexOf(
                    "dash") >= 0)
            {
                return DashStyle.Dash;
            }

            if (lower.IndexOf(
                    "dot") >= 0)
            {
                return DashStyle.Dot;
            }

            return DashStyle.Solid;
        }

        private static void DrawChartMarker(
            Graphics g,
            PointF point,
            string symbol,
            float size,
            Color color)
        {
            size =
                Math.Max(
                    3f,
                    Math.Min(
                        24f,
                        size));

            float radius =
                size /
                2f;

            RectangleF markerRect =
                new RectangleF(
                    point.X - radius,
                    point.Y - radius,
                    size,
                    size);

            string normalized =
                string.IsNullOrEmpty(symbol)
                    ? "circle"
                    : symbol.ToLowerInvariant();

            using (Brush brush =
                new SolidBrush(color))
            using (Pen pen =
                new Pen(
                    color,
                    Math.Max(
                        1f,
                        size /
                        7f)))
            {
                if (normalized == "square")
                {
                    g.FillRectangle(
                        brush,
                        markerRect);
                }
                else if (normalized == "diamond")
                {
                    PointF[] diamond =
                        new PointF[]
                        {
                            new PointF(
                                point.X,
                                markerRect.Top),
                            new PointF(
                                markerRect.Right,
                                point.Y),
                            new PointF(
                                point.X,
                                markerRect.Bottom),
                            new PointF(
                                markerRect.Left,
                                point.Y)
                        };

                    g.FillPolygon(
                        brush,
                        diamond);
                }
                else if (normalized == "triangle")
                {
                    PointF[] triangle =
                        new PointF[]
                        {
                            new PointF(
                                point.X,
                                markerRect.Top),
                            new PointF(
                                markerRect.Right,
                                markerRect.Bottom),
                            new PointF(
                                markerRect.Left,
                                markerRect.Bottom)
                        };

                    g.FillPolygon(
                        brush,
                        triangle);
                }
                else if (normalized == "x")
                {
                    g.DrawLine(
                        pen,
                        markerRect.Left,
                        markerRect.Top,
                        markerRect.Right,
                        markerRect.Bottom);

                    g.DrawLine(
                        pen,
                        markerRect.Right,
                        markerRect.Top,
                        markerRect.Left,
                        markerRect.Bottom);
                }
                else if (normalized == "plus")
                {
                    g.DrawLine(
                        pen,
                        point.X,
                        markerRect.Top,
                        point.X,
                        markerRect.Bottom);

                    g.DrawLine(
                        pen,
                        markerRect.Left,
                        point.Y,
                        markerRect.Right,
                        point.Y);
                }
                else
                {
                    g.FillEllipse(
                        brush,
                        markerRect);
                }
            }
        }

        private static ChartLineStyle ReadChartMinorGridlineStyle(
            XmlDocument chartDoc,
            Dictionary<string, Color> theme)
        {
            XmlNode valueAxis =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "valAx");

            XmlNode minorGridlines =
                valueAxis == null
                    ? null
                    : DirectChild(
                        valueAxis,
                        "minorGridlines");

            if (minorGridlines == null)
                return null;

            ChartLineStyle style =
                new ChartLineStyle();
            style.Color =
                Color.FromArgb(
                    238,
                    240,
                    243);
            style.Width = 1f;

            XmlNode shapeProperties =
                DirectChild(
                    minorGridlines,
                    "spPr");
            XmlNode line =
                shapeProperties == null
                    ? null
                    : DirectChild(
                        shapeProperties,
                        "ln");

            if (line == null)
                return style;

            Color? color =
                ReadSolidFill(
                    line,
                    theme);

            if (color.HasValue)
            {
                style.Color =
                    color.Value;
            }

            long width =
                GetLong(
                    line,
                    "w",
                    0);

            if (width > 0)
            {
                style.Width =
                    Math.Max(
                        1f,
                        EmuToRenderPixels(
                            width));
            }

            XmlNode dash =
                DirectChild(
                    line,
                    "prstDash");

            style.DashStyle =
                ParseChartDashStyle(
                    dash == null
                        ? string.Empty
                        : GetAttr(
                            dash,
                            "val"));

            return style;
        }

        private static ChartLineStyle ReadChartLeaderLineStyle(
            XmlDocument chartDoc,
            Dictionary<string, Color> theme)
        {
            ChartLineStyle style =
                new ChartLineStyle();
            style.Color =
                Color.FromArgb(
                    120,
                    120,
                    120);
            style.Width = 1f;

            XmlNode labels =
                FindChartLevelDataLabels(
                    chartDoc);
            XmlNode leaderLines =
                labels == null
                    ? null
                    : DirectChild(
                        labels,
                        "leaderLines");
            XmlNode shapeProperties =
                leaderLines == null
                    ? null
                    : DirectChild(
                        leaderLines,
                        "spPr");
            XmlNode line =
                shapeProperties == null
                    ? null
                    : DirectChild(
                        shapeProperties,
                        "ln");

            if (line == null)
                return style;

            Color? color =
                ReadSolidFill(
                    line,
                    theme);

            if (color.HasValue)
            {
                style.Color =
                    color.Value;
            }

            long width =
                GetLong(
                    line,
                    "w",
                    0);

            if (width > 0)
            {
                style.Width =
                    Math.Max(
                        1f,
                        EmuToRenderPixels(
                            width));
            }

            XmlNode dash =
                DirectChild(
                    line,
                    "prstDash");

            style.DashStyle =
                ParseChartDashStyle(
                    dash == null
                        ? string.Empty
                        : GetAttr(
                            dash,
                            "val"));

            return style;
        }

        private static ChartLineStyle ReadChartMajorGridlineStyle(
            XmlDocument chartDoc,
            Dictionary<string, Color> theme)
        {
            ChartLineStyle style =
                new ChartLineStyle();
            style.Color =
                Color.FromArgb(
                    225,
                    228,
                    232);
            style.Width = 1f;

            XmlNode valueAxis =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "valAx");

            XmlNode majorGridlines =
                valueAxis == null
                    ? null
                    : DirectChild(
                        valueAxis,
                        "majorGridlines");
            XmlNode shapeProperties =
                majorGridlines == null
                    ? null
                    : DirectChild(
                        majorGridlines,
                        "spPr");
            XmlNode line =
                shapeProperties == null
                    ? null
                    : DirectChild(
                        shapeProperties,
                        "ln");

            if (line == null)
                return style;

            Color? color =
                ReadSolidFill(
                    line,
                    theme);

            if (color.HasValue)
            {
                style.Color =
                    color.Value;
            }

            long width =
                GetLong(
                    line,
                    "w",
                    0);

            if (width > 0)
            {
                style.Width =
                    Math.Max(
                        1f,
                        EmuToRenderPixels(
                            width));
            }

            XmlNode dash =
                DirectChild(
                    line,
                    "prstDash");

            style.DashStyle =
                ParseChartDashStyle(
                    dash == null
                        ? string.Empty
                        : GetAttr(
                            dash,
                            "val"));

            return style;
        }

        private static ChartLineStyle ReadChartAxisLineStyle(
            XmlDocument chartDoc,
            string axisName,
            Dictionary<string, Color> theme)
        {
            ChartLineStyle style =
                new ChartLineStyle();

            if (chartDoc == null ||
                string.IsNullOrEmpty(
                    axisName))
            {
                return style;
            }

            XmlNode axis =
                FindFirst(
                    chartDoc,
                    axisName);

            if (axis == null)
                return style;

            XmlNode shapeProperties =
                DirectChild(
                    axis,
                    "spPr");
            XmlNode line =
                shapeProperties == null
                    ? null
                    : DirectChild(
                        shapeProperties,
                        "ln");

            if (line == null)
                return style;

            Color? color =
                ReadSolidFill(
                    line,
                    theme);

            if (color.HasValue)
            {
                style.Color =
                    color.Value;
            }

            long width =
                GetLong(
                    line,
                    "w",
                    0);

            if (width > 0)
            {
                style.Width =
                    Math.Max(
                        1f,
                        EmuToRenderPixels(
                            width));
            }

            XmlNode dash =
                DirectChild(
                    line,
                    "prstDash");

            style.DashStyle =
                ParseChartDashStyle(
                    dash == null
                        ? string.Empty
                        : GetAttr(
                            dash,
                            "val"));

            return style;
        }

        private static double ReadChartCategoryAxisCrossValue(
            XmlDocument chartDoc,
            ChartAxisScale scale)
        {
            if (scale == null)
                return 0.0;

            XmlNode categoryAxis =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "catAx");

            if (categoryAxis == null)
                return 0.0;

            XmlNode crossesAt =
                DirectChild(
                    categoryAxis,
                    "crossesAt");

            double parsed;

            if (crossesAt != null &&
                double.TryParse(
                    GetAttr(
                        crossesAt,
                        "val"),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out parsed) &&
                !double.IsNaN(parsed) &&
                !double.IsInfinity(parsed))
            {
                return parsed;
            }

            XmlNode crosses =
                DirectChild(
                    categoryAxis,
                    "crosses");

            string value =
                crosses == null
                    ? string.Empty
                    : GetAttr(
                        crosses,
                        "val") ??
                      string.Empty;

            if (string.Equals(
                    value,
                    "max",
                    StringComparison.OrdinalIgnoreCase))
            {
                return scale.Maximum;
            }

            if (string.Equals(
                    value,
                    "min",
                    StringComparison.OrdinalIgnoreCase))
            {
                return scale.Minimum;
            }

            return 0.0;
        }

        private static string ReadChartAxisTickLabelPosition(
            XmlDocument chartDoc,
            string axisName)
        {
            if (chartDoc == null ||
                string.IsNullOrEmpty(
                    axisName))
            {
                return string.Empty;
            }

            XmlNode axis =
                FindFirst(
                    chartDoc,
                    axisName);

            if (axis == null)
                return string.Empty;

            XmlNode tickLabelPosition =
                DirectChild(
                    axis,
                    "tickLblPos");

            return tickLabelPosition == null
                ? string.Empty
                : GetAttr(
                    tickLabelPosition,
                    "val") ??
                  string.Empty;
        }

        private static int ReadChartCategoryLabelSkip(
            XmlDocument chartDoc)
        {
            XmlNode categoryAxis =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "catAx");

            if (categoryAxis == null)
                return 1;

            XmlNode skip =
                DirectChild(
                    categoryAxis,
                    "tickLblSkip");

            int parsed;

            if (skip != null &&
                int.TryParse(
                    GetAttr(
                        skip,
                        "val"),
                    out parsed) &&
                parsed > 0)
            {
                return Math.Min(
                    1000,
                    parsed);
            }

            return 1;
        }

        private static string ReadChartAxisTitle(
            XmlDocument chartDoc,
            string axisName)
        {
            if (chartDoc == null ||
                string.IsNullOrEmpty(axisName))
            {
                return string.Empty;
            }

            XmlNode axis =
                FindFirst(
                    chartDoc,
                    axisName);

            if (axis == null)
                return string.Empty;

            XmlNode title =
                DirectChild(
                    axis,
                    "title");

            if (title == null)
                return string.Empty;

            return ReadChartTitleText(
                title);
        }

        private static string ReadChartTitleText(
            XmlNode title)
        {
            if (title == null)
                return string.Empty;

            StringBuilder text =
                new StringBuilder();

            foreach (XmlNode node in
                FindAll(
                    title,
                    "t"))
            {
                text.Append(
                    node.InnerText);
            }

            if (text.Length == 0)
            {
                foreach (XmlNode node in
                    FindAll(
                        title,
                        "v"))
                {
                    text.Append(
                        node.InnerText);
                }
            }

            return text.ToString();
        }

        private static void DrawChartAxisTitles(
            Graphics g,
            RectangleF plot,
            RectangleF chartRect,
            string kind,
            string valueTitle,
            string categoryTitle)
        {
            using (Font font =
                SafeFont(
                    "Arial",
                    Math.Max(
                        7f,
                        Math.Min(
                            11f,
                            chartRect.Height /
                            30f)),
                    FontStyle.Bold))
            using (Brush brush =
                new SolidBrush(
                    Color.FromArgb(
                        70,
                        70,
                        70)))
            using (StringFormat format =
                new StringFormat())
            {
                format.Alignment =
                    StringAlignment.Center;
                format.LineAlignment =
                    StringAlignment.Center;
                format.Trimming =
                    StringTrimming.EllipsisCharacter;

                if (kind == "bar")
                {
                    if (!string.IsNullOrEmpty(
                            valueTitle))
                    {
                        g.DrawString(
                            valueTitle,
                            font,
                            brush,
                            new RectangleF(
                                plot.Left,
                                plot.Bottom + 24f,
                                plot.Width,
                                Math.Max(
                                    18f,
                                    font.Height + 4f)),
                            format);
                    }

                    if (!string.IsNullOrEmpty(
                            categoryTitle))
                    {
                        DrawRotatedChartAxisTitle(
                            g,
                            categoryTitle,
                            font,
                            brush,
                            new PointF(
                                chartRect.Left + 12f,
                                plot.Top +
                                plot.Height /
                                2f),
                            -90f);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(
                            categoryTitle))
                    {
                        g.DrawString(
                            categoryTitle,
                            font,
                            brush,
                            new RectangleF(
                                plot.Left,
                                plot.Bottom + 24f,
                                plot.Width,
                                Math.Max(
                                    18f,
                                    font.Height + 4f)),
                            format);
                    }

                    if (!string.IsNullOrEmpty(
                            valueTitle))
                    {
                        DrawRotatedChartAxisTitle(
                            g,
                            valueTitle,
                            font,
                            brush,
                            new PointF(
                                chartRect.Left + 12f,
                                plot.Top +
                                plot.Height /
                                2f),
                            -90f);
                    }
                }
            }
        }

        private static void DrawRotatedChartAxisTitle(
            Graphics g,
            string text,
            Font font,
            Brush brush,
            PointF center,
            float degrees)
        {
            GraphicsState state =
                g.Save();

            try
            {
                g.TranslateTransform(
                    center.X,
                    center.Y);
                g.RotateTransform(
                    degrees);

                SizeF size =
                    g.MeasureString(
                        text,
                        font);

                g.DrawString(
                    text,
                    font,
                    brush,
                    -size.Width /
                        2f,
                    -size.Height /
                        2f);
            }
            finally
            {
                g.Restore(state);
            }
        }

        private static string FormatChartAxisNumber(
            double value,
            string formatCode)
        {
            if (string.IsNullOrEmpty(
                    formatCode) ||
                string.Equals(
                    formatCode,
                    "General",
                    StringComparison.OrdinalIgnoreCase))
            {
                return FormatChartNumber(
                    value);
            }

            string code =
                formatCode;

            int section =
                code.IndexOf(';');

            if (section >= 0)
            {
                code =
                    code.Substring(
                        0,
                        section);
            }

            if (code.IndexOf(
                    '%') >= 0)
            {
                int decimals =
                    CountChartFormatDecimals(
                        code);

                return value.ToString(
                    "P" +
                    decimals.ToString(),
                    System.Globalization.CultureInfo.InvariantCulture);
            }

            bool thousands =
                code.IndexOf(
                    ',') >= 0;

            int decimalPlaces =
                CountChartFormatDecimals(
                    code);

            string numeric =
                thousands
                    ? "N" +
                        decimalPlaces.ToString()
                    : "F" +
                        decimalPlaces.ToString();

            return value.ToString(
                numeric,
                System.Globalization.CultureInfo.InvariantCulture);
        }

        private static int CountChartFormatDecimals(
            string formatCode)
        {
            if (string.IsNullOrEmpty(
                    formatCode))
            {
                return 0;
            }

            int dot =
                formatCode.IndexOf('.');

            if (dot < 0)
                return 0;

            int count = 0;

            for (int i = dot + 1;
                 i < formatCode.Length;
                 i++)
            {
                char ch =
                    formatCode[i];

                if (ch == '0' ||
                    ch == '#')
                {
                    count++;
                }
                else
                {
                    break;
                }
            }

            return Math.Max(
                0,
                Math.Min(
                    6,
                    count));
        }

        private static string ReadChartSeriesName(XmlNode ser)
        {
            XmlNode tx = DirectChild(ser, "tx");
            if (tx == null) return "";

            XmlNode v = FindFirst(tx, "v");
            return v != null ? v.InnerText : "";
        }

        private static void ReadChartCategories(XmlNode ser, List<string> output)
        {
            XmlNode cat = DirectChild(ser, "cat");
            if (cat == null) return;

            XmlNode cache = FindFirst(cat, "strCache");
            if (cache == null) cache = FindFirst(cat, "strLit");
            if (cache == null) cache = FindFirst(cat, "numCache");
            if (cache == null) cache = FindFirst(cat, "numLit");
            if (cache == null) return;

            foreach (XmlNode pt in FindAll(cache, "pt"))
            {
                XmlNode v = FindFirst(pt, "v");
                if (v != null) output.Add(v.InnerText);
            }
        }

        private static void ReadChartValues(XmlNode ser, List<double> output)
        {
            XmlNode val = DirectChild(ser, "val");
            if (val == null) val = DirectChild(ser, "yVal");
            if (val == null) return;

            XmlNode cache = FindFirst(val, "numCache");
            if (cache == null) cache = FindFirst(val, "numLit");
            if (cache == null) return;

            foreach (XmlNode pt in FindAll(cache, "pt"))
            {
                XmlNode v = FindFirst(pt, "v");
                double parsed;

                if (v != null && double.TryParse(
                    v.InnerText,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out parsed))
                {
                    output.Add(parsed);
                }
            }
        }

        private static string ReadChartTitle(XmlDocument chartDoc)
        {
            XmlNode title =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "title");

            return ReadChartTitleText(
                title);
        }

        private static void DrawChartLegend(
            Graphics g,
            RectangleF rect,
            List<ChartSeriesData> series,
            Color[] palette,
            string kind,
            string position)
        {
            if (string.IsNullOrEmpty(position) ||
                series == null ||
                series.Count == 0 ||
                palette == null ||
                palette.Length == 0)
            {
                return;
            }

            List<string> labels =
                new List<string>();

            List<Color> colors =
                new List<Color>();

            if (kind == "pie" &&
                series[0].Categories.Count > 0)
            {
                for (int i = 0;
                     i < series[0].Categories.Count;
                     i++)
                {
                    labels.Add(
                        string.IsNullOrEmpty(
                            series[0].Categories[i])
                            ? "Item " +
                                (i + 1).ToString()
                            : series[0].Categories[i]);

                    colors.Add(
                        GetChartPointColor(
                            series[0],
                            i,
                            palette));
                }
            }
            else
            {
                for (int i = 0;
                     i < series.Count;
                     i++)
                {
                    labels.Add(
                        string.IsNullOrEmpty(
                            series[i].Name)
                            ? "Series " +
                                (i + 1).ToString()
                            : series[i].Name);

                    colors.Add(
                        GetChartSeriesColor(
                            series,
                            i,
                            palette));
                }
            }

            if (labels.Count == 0)
                return;

            using (Font font = SafeFont(
                "Arial",
                Math.Max(
                    7f,
                    Math.Min(
                        11f,
                        rect.Height /
                        28f))))
            using (Brush text =
                new SolidBrush(
                    Color.FromArgb(
                        60,
                        60,
                        60)))
            {
                if (position == "t" ||
                    position == "b")
                {
                    float y =
                        position == "t"
                            ? rect.Top +
                                Math.Max(
                                    28f,
                                    rect.Height *
                                    0.08f)
                            : rect.Bottom -
                                Math.Max(
                                    28f,
                                    rect.Height *
                                    0.08f);

                    float x =
                        rect.Left +
                        12f;

                    for (int i = 0;
                         i < labels.Count;
                         i++)
                    {
                        SizeF measured =
                            g.MeasureString(
                                labels[i],
                                font);

                        float itemWidth =
                            16f +
                            measured.Width +
                            12f;

                        if (x + itemWidth >
                            rect.Right - 8f)
                        {
                            x =
                                rect.Left +
                                12f;
                            y +=
                                font.Height +
                                7f;
                        }

                        using (Brush swatch =
                            new SolidBrush(
                                colors[i]))
                        {
                            g.FillRectangle(
                                swatch,
                                x,
                                y + 3f,
                                10f,
                                10f);
                        }

                        g.DrawString(
                            labels[i],
                            font,
                            text,
                            x + 14f,
                            y);

                        x +=
                            itemWidth;
                    }

                    return;
                }

                float legendWidth =
                    Math.Max(
                        72f,
                        rect.Width *
                        0.17f);

                float legendX =
                    position == "l"
                        ? rect.Left + 10f
                        : rect.Right -
                            legendWidth;

                float legendY =
                    position == "tr"
                        ? rect.Top + 10f
                        : rect.Top +
                            Math.Max(
                                35f,
                                rect.Height *
                                0.12f);

                for (int i = 0;
                     i < labels.Count;
                     i++)
                {
                    using (Brush swatch =
                        new SolidBrush(
                            colors[i]))
                    {
                        g.FillRectangle(
                            swatch,
                            legendX,
                            legendY + 3f,
                            10f,
                            10f);
                    }

                    RectangleF labelRect =
                        new RectangleF(
                            legendX + 14f,
                            legendY,
                            Math.Max(
                                20f,
                                legendWidth -
                                18f),
                            font.Height +
                            3f);

                    using (StringFormat format =
                        new StringFormat())
                    {
                        format.Trimming =
                            StringTrimming.EllipsisCharacter;
                        format.FormatFlags =
                            StringFormatFlags.NoWrap;

                        g.DrawString(
                            labels[i],
                            font,
                            text,
                            labelRect,
                            format);
                    }

                    legendY +=
                        font.Height +
                        4f;
                }
            }
        }

        private static void DrawPieChart(
            Graphics g,
            RectangleF plot,
            ChartSeriesData series,
            Color[] palette,
            float startAngle)
        {
            double total = 0;
            for (int i = 0; i < series.Values.Count; i++)
                total += Math.Abs(series.Values[i]);

            if (total <= 0) return;

            float diameter = Math.Min(plot.Width, plot.Height) * 0.82f;
            RectangleF pie = new RectangleF(
                plot.Left + (plot.Width - diameter) / 2f,
                plot.Top + (plot.Height - diameter) / 2f,
                diameter,
                diameter);

            float start = startAngle;

            for (int i = 0; i < series.Values.Count; i++)
            {
                float sweep = (float)(360.0 * Math.Abs(series.Values[i]) / total);
                Color sliceColor =
                    GetChartPointColor(
                        series,
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
                        start,
                        sweep);
                }
                start += sweep;
            }
        }

        private static void DrawSmartArt(
            Graphics g,
            XmlDocument dataDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            List<string> labels = new List<string>();

            foreach (XmlNode pt in FindAll(dataDoc, "pt"))
            {
                string type = GetAttr(pt, "type");
                if (!string.IsNullOrEmpty(type) && type != "node" && type != "asst")
                    continue;

                StringBuilder sb = new StringBuilder();
                foreach (XmlNode t in FindAll(pt, "t"))
                    sb.Append(t.InnerText);

                string label = sb.ToString().Trim();
                if (!string.IsNullOrEmpty(label) && !labels.Contains(label))
                    labels.Add(label);
            }

            if (labels.Count == 0)
            {
                DrawPlaceholder(g, rect, "SmartArt / Diagram");
                return;
            }

            int count = labels.Count;
            int cols = count <= 3 ? count : (int)Math.Ceiling(Math.Sqrt(count));
            int rows = (int)Math.Ceiling(count / (double)cols);

            float gap = Math.Max(8f, Math.Min(rect.Width, rect.Height) * 0.035f);
            float cellW = Math.Max(20f, (rect.Width - gap * (cols + 1)) / cols);
            float cellH = Math.Max(20f, (rect.Height - gap * (rows + 1)) / rows);

            Color fill = ThemeOrDefault(theme, "accent1", Color.FromArgb(79,129,189));
            Color line = ControlPaint.Dark(fill);

            for (int i = 0; i < labels.Count; i++)
            {
                int row = i / cols;
                int col = i % cols;
                RectangleF box = new RectangleF(
                    rect.Left + gap + col * (cellW + gap),
                    rect.Top + gap + row * (cellH + gap),
                    cellW,
                    cellH);

                using (GraphicsPath path = BuildPresetPath("roundRect", box))
                using (Brush b = new SolidBrush(Color.FromArgb(235, fill)))
                using (Pen p = new Pen(line, 1.5f))
                {
                    g.FillPath(b, path);
                    g.DrawPath(p, path);
                }

                using (Font font = SafeFont("Arial", Math.Max(8f, Math.Min(14f, cellH / 5f))))
                using (Brush tb = new SolidBrush(Color.White))
                using (StringFormat sf = new StringFormat())
                {
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Trimming = StringTrimming.EllipsisCharacter;
                    g.DrawString(labels[i], font, tb, box, sf);
                }

                if (i > 0)
                {
                    int prevRow = (i - 1) / cols;
                    int prevCol = (i - 1) % cols;
                    PointF from = new PointF(
                        rect.Left + gap + prevCol * (cellW + gap) + cellW / 2f,
                        rect.Top + gap + prevRow * (cellH + gap) + cellH / 2f);
                    PointF to = new PointF(box.Left + box.Width / 2f, box.Top + box.Height / 2f);

                    using (Pen connector = new Pen(Color.FromArgb(100, line), 1.2f))
                        g.DrawLine(connector, from, to);
                }
            }
        }
    }
}
