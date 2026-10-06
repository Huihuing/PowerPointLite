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
            public bool Reverse;
        }

        private sealed class ChartLabelOptions
        {
            public bool ShowValue;
            public bool ShowCategoryName;
            public bool ShowSeriesName;
            public bool ShowPercent;
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

                if (data.Values.Count > 0)
                    series.Add(data);
            }

            using (Brush bg =
                new SolidBrush(Color.White))
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
                    chartDoc);

            if (kind == "pie")
            {
                DrawPieChart(
                    g,
                    plot,
                    series[0],
                    palette);

                if (labelOptions.HasAny)
                {
                    DrawPieChartValueLabels(
                        g,
                        plot,
                        series[0],
                        labelOptions);
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

            minValue =
                axisScale.Minimum;
            maxValue =
                axisScale.Maximum;

            DrawChartValueGrid(
                g,
                plot,
                axisScale,
                kind);

            double range =
                Math.Max(
                    0.0000001,
                    maxValue - minValue);

            float zeroY =
                plot.Bottom -
                (float)(
                    ChartAxisFraction(
                        0.0,
                        axisScale) *
                    plot.Height);

            float zeroX =
                plot.Left +
                (float)(
                    ChartAxisFraction(
                        0.0,
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

            using (Pen axis =
                new Pen(
                    Color.FromArgb(
                        100,
                        100,
                        100),
                    1.2f))
            {
                if (kind == "bar")
                {
                    g.DrawLine(
                        axis,
                        zeroX,
                        plot.Top,
                        zeroX,
                        plot.Bottom);

                    g.DrawLine(
                        axis,
                        plot.Left,
                        plot.Bottom,
                        plot.Right,
                        plot.Bottom);
                }
                else
                {
                    g.DrawLine(
                        axis,
                        plot.Left,
                        zeroY,
                        plot.Right,
                        zeroY);

                    g.DrawLine(
                        axis,
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
                            palette[
                                si %
                                palette.Length];

                        using (Pen pen =
                            new Pen(
                                color,
                                Math.Max(
                                    2f,
                                    plot.Width /
                                    250f)))
                        using (Brush marker =
                            new SolidBrush(color))
                        {
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

                                g.FillEllipse(
                                    marker,
                                    point.X - 3,
                                    point.Y - 3,
                                    6,
                                    6);

                                if (labelOptions.HasAny &&
                                    i < sd.Values.Count)
                                {
                                    string label =
                                        BuildChartDataLabel(
                                            labelOptions,
                                            sd,
                                            i,
                                            false);

                                    DrawChartValueLabel(
                                        g,
                                        valueFont,
                                        valueBrush,
                                        label,
                                        point.X,
                                        point.Y - 16f,
                                        true);
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

                    float barH =
                        Math.Max(
                            2f,
                            groupH *
                            0.75f /
                            series.Count);

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

                            float valueX =
                                plot.Left +
                                (float)(
                                    ChartAxisFraction(
                                        value,
                                        axisScale) *
                                    plot.Width);

                            float left =
                                Math.Min(
                                    zeroX,
                                    valueX);

                            float width =
                                Math.Max(
                                    1f,
                                    Math.Abs(
                                        valueX -
                                        zeroX));

                            float y =
                                plot.Top +
                                ci * groupH +
                                groupH * 0.12f +
                                si * barH;

                            Color color =
                                palette[
                                    si %
                                    palette.Length];

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

                            if (labelOptions.HasAny)
                            {
                                float labelX =
                                    value >= 0.0
                                        ? valueX + 4f
                                        : valueX - 4f;

                                string label =
                                    BuildChartDataLabel(
                                        labelOptions,
                                        series[si],
                                        ci,
                                        false);

                                DrawChartValueLabel(
                                    g,
                                    valueFont,
                                    valueBrush,
                                    label,
                                    labelX,
                                    y +
                                        Math.Max(
                                            0f,
                                            (barH -
                                             valueFont.Height) /
                                            2f),
                                    value < 0.0);
                            }
                        }
                    }
                }
                else
                {
                    float groupW =
                        plot.Width /
                        categoryCount;

                    float barW =
                        Math.Max(
                            2f,
                            groupW *
                            0.75f /
                            series.Count);

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

                            float valueY =
                                plot.Bottom -
                                (float)(
                                    ChartAxisFraction(
                                        value,
                                        axisScale) *
                                    plot.Height);

                            float top =
                                Math.Min(
                                    zeroY,
                                    valueY);

                            float height =
                                Math.Max(
                                    1f,
                                    Math.Abs(
                                        valueY -
                                        zeroY));

                            float x =
                                plot.Left +
                                ci * groupW +
                                groupW * 0.12f +
                                si * barW;

                            Color color =
                                palette[
                                    si %
                                    palette.Length];

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

                            if (labelOptions.HasAny)
                            {
                                float labelY =
                                    value >= 0.0
                                        ? valueY -
                                            valueFont.Height -
                                            2f
                                        : valueY + 2f;

                                string label =
                                    BuildChartDataLabel(
                                        labelOptions,
                                        series[si],
                                        ci,
                                        false);

                                DrawChartValueLabel(
                                    g,
                                    valueFont,
                                    valueBrush,
                                    label,
                                    x +
                                        Math.Max(
                                            1f,
                                            (barW - 1f) /
                                            2f),
                                    labelY,
                                    true);
                            }
                        }
                    }
                }
            }

            DrawChartCategoryLabels(
                g,
                plot,
                series[0],
                categoryCount,
                kind);

            DrawChartLegend(
                g,
                rect,
                series,
                palette,
                kind,
                legendPosition);
        }

        private static ChartLabelOptions ReadChartLabelOptions(
            XmlDocument chartDoc)
        {
            ChartLabelOptions result =
                new ChartLabelOptions();

            XmlNode labels =
                chartDoc == null
                    ? null
                    : FindFirst(
                        chartDoc,
                        "dLbls");

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
                    FormatChartNumber(
                        series.Values[index]));
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
            string kind)
        {
            if (scale == null)
                return;

            List<double> ticks =
                BuildChartAxisTicks(
                    scale);

            using (Pen grid =
                new Pen(
                    Color.FromArgb(
                        225,
                        228,
                        232),
                    1f))
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

                        string label =
                            FormatChartNumber(
                                value);

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
                            plot.Bottom + 2f);
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

                        string label =
                            FormatChartNumber(
                                value);

                        SizeF size =
                            g.MeasureString(
                                label,
                                font);

                        g.DrawString(
                            label,
                            font,
                            text,
                            plot.Left -
                                size.Width -
                                4f,
                            y -
                                size.Height /
                                2f);
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
            ChartLabelOptions options)
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

            float radius =
                diameter *
                0.34f;

            float cx =
                plot.Left +
                plot.Width /
                2f;

            float cy =
                plot.Top +
                plot.Height /
                2f;

            float start =
                -90f;

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

                    string label =
                        BuildChartDataLabel(
                            options,
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

                    g.DrawString(
                        label,
                        font,
                        brush,
                        x,
                        y);

                    start += sweep;
                }
            }
        }

        private static void DrawChartCategoryLabels(
            Graphics g,
            RectangleF plot,
            ChartSeriesData firstSeries,
            int categoryCount,
            string kind)
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

                for (int i = 0; i < count; i++)
                {
                    if (kind == "bar")
                    {
                        sf.Alignment = StringAlignment.Far;
                        float groupH = plot.Height / Math.Max(1, categoryCount);
                        RectangleF label = new RectangleF(
                            plot.Left - Math.Max(24f, plot.Width * 0.09f),
                            plot.Top + i * groupH,
                            Math.Max(20f, plot.Width * 0.08f),
                            groupH);
                        sf.LineAlignment = StringAlignment.Center;
                        g.DrawString(firstSeries.Categories[i], font, brush, label, sf);
                    }
                    else
                    {
                        sf.Alignment = StringAlignment.Center;
                        sf.LineAlignment = StringAlignment.Near;
                        float groupW = plot.Width / Math.Max(1, categoryCount);
                        RectangleF label = new RectangleF(
                            plot.Left + i * groupW,
                            plot.Bottom + 2f,
                            groupW,
                            Math.Max(14f, font.Height + 4f));
                        g.DrawString(firstSeries.Categories[i], font, brush, label, sf);
                    }
                }
            }
        }

        private static Color ThemeOrDefault(
            Dictionary<string, Color> theme,
            string key,
            Color fallback)
        {
            return theme != null && theme.ContainsKey(key) ? theme[key] : fallback;
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
            XmlNode title = FindFirst(chartDoc, "title");
            if (title == null) return "";

            StringBuilder sb = new StringBuilder();
            foreach (XmlNode t in FindAll(title, "t")) sb.Append(t.InnerText);
            if (sb.Length == 0)
                foreach (XmlNode v in FindAll(title, "v")) sb.Append(v.InnerText);
            return sb.ToString();
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
                        palette[
                            i %
                            palette.Length]);
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
                        palette[
                            i %
                            palette.Length]);
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

                float x =
                    position == "l"
                        ? rect.Left + 10f
                        : rect.Right -
                            legendWidth;

                float y =
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
                            x,
                            y + 3f,
                            10f,
                            10f);
                    }

                    RectangleF labelRect =
                        new RectangleF(
                            x + 14f,
                            y,
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

                    y +=
                        font.Height +
                        4f;
                }
            }
        }

        private static void DrawPieChart(
            Graphics g,
            RectangleF plot,
            ChartSeriesData series,
            Color[] palette)
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

            float start = -90f;

            for (int i = 0; i < series.Values.Count; i++)
            {
                float sweep = (float)(360.0 * Math.Abs(series.Values[i]) / total);
                using (Brush brush = new SolidBrush(palette[i % palette.Length]))
                    g.FillPie(brush, pie.X, pie.Y, pie.Width, pie.Height, start, sweep);
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
