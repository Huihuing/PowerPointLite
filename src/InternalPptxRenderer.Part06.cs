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
        private static void DrawChart(
            Graphics g,
            XmlDocument chartDoc,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (chartDoc == null)
                return;

            string kind = "column";

            if (FindFirst(chartDoc, "pieChart") != null || FindFirst(chartDoc, "doughnutChart") != null)
                kind = "pie";
            else if (FindFirst(chartDoc, "lineChart") != null)
                kind = "line";
            else if (FindFirst(chartDoc, "barChart") != null)
            {
                XmlNode barDir = FindFirst(chartDoc, "barDir");
                kind = barDir != null && GetAttr(barDir, "val") == "bar" ? "bar" : "column";
            }
            else if (FindFirst(chartDoc, "areaChart") != null)
                kind = "line";

            List<ChartSeriesData> series = new List<ChartSeriesData>();

            foreach (XmlNode ser in FindAll(chartDoc, "ser"))
            {
                ChartSeriesData data = new ChartSeriesData();
                data.Name = ReadChartSeriesName(ser);
                ReadChartCategories(ser, data.Categories);
                ReadChartValues(ser, data.Values);

                if (data.Values.Count > 0)
                    series.Add(data);
            }

            using (Brush bg = new SolidBrush(Color.White))
                g.FillRectangle(bg, rect);

            using (Pen border = new Pen(Color.FromArgb(210, 210, 210), 1f))
                g.DrawRectangle(border, rect.X, rect.Y, rect.Width, rect.Height);

            if (series.Count == 0)
            {
                DrawPlaceholder(g, rect, "Chart (no cached data)");
                return;
            }

            string title = ReadChartTitle(chartDoc);
            float topPad = string.IsNullOrEmpty(title) ? 12f : Math.Max(34f, rect.Height * 0.10f);

            if (!string.IsNullOrEmpty(title))
            {
                using (Font font = SafeFont("Arial", Math.Max(10f, Math.Min(18f, rect.Height / 18f)), FontStyle.Bold))
                using (Brush brush = new SolidBrush(Color.FromArgb(45, 45, 45)))
                using (StringFormat sf = new StringFormat())
                {
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    g.DrawString(title, font, brush,
                        new RectangleF(rect.Left + 5, rect.Top + 4, rect.Width - 10, topPad - 4), sf);
                }
            }

            RectangleF plot = new RectangleF(
                rect.Left + Math.Max(28f, rect.Width * 0.08f),
                rect.Top + topPad,
                Math.Max(10f, rect.Width - Math.Max(70f, rect.Width * 0.18f)),
                Math.Max(10f, rect.Height - topPad - Math.Max(30f, rect.Height * 0.10f)));

            Color[] palette = new Color[]
            {
                ThemeOrDefault(theme, "accent1", Color.FromArgb(79,129,189)),
                ThemeOrDefault(theme, "accent2", Color.FromArgb(192,80,77)),
                ThemeOrDefault(theme, "accent3", Color.FromArgb(155,187,89)),
                ThemeOrDefault(theme, "accent4", Color.FromArgb(128,100,162)),
                ThemeOrDefault(theme, "accent5", Color.FromArgb(75,172,198)),
                ThemeOrDefault(theme, "accent6", Color.FromArgb(247,150,70))
            };

            if (kind == "pie")
            {
                DrawPieChart(g, plot, series[0], palette);
                return;
            }

            double max = 0;
            int categoryCount = 0;

            for (int sIndex = 0; sIndex < series.Count; sIndex++)
            {
                categoryCount = Math.Max(categoryCount, series[sIndex].Values.Count);
                for (int i = 0; i < series[sIndex].Values.Count; i++)
                    max = Math.Max(max, Math.Abs(series[sIndex].Values[i]));
            }

            if (max <= 0) max = 1;
            if (categoryCount <= 0) categoryCount = 1;

            using (Pen axis = new Pen(Color.FromArgb(100, 100, 100), 1f))
            {
                g.DrawLine(axis, plot.Left, plot.Bottom, plot.Right, plot.Bottom);
                g.DrawLine(axis, plot.Left, plot.Top, plot.Left, plot.Bottom);
            }

            if (kind == "line")
            {
                for (int si = 0; si < series.Count; si++)
                {
                    ChartSeriesData sd = series[si];
                    if (sd.Values.Count == 0) continue;

                    List<PointF> points = new List<PointF>();

                    for (int i = 0; i < sd.Values.Count; i++)
                    {
                        float x = categoryCount <= 1
                            ? plot.Left + plot.Width / 2f
                            : plot.Left + plot.Width * i / (categoryCount - 1f);
                        float y = plot.Bottom - (float)(plot.Height * sd.Values[i] / max);
                        points.Add(new PointF(x, y));
                    }

                    using (Pen pen = new Pen(palette[si % palette.Length], Math.Max(2f, plot.Width / 250f)))
                    using (Brush marker = new SolidBrush(palette[si % palette.Length]))
                    {
                        if (points.Count > 1)
                            g.DrawLines(pen, points.ToArray());

                        foreach (PointF point in points)
                            g.FillEllipse(marker, point.X - 3, point.Y - 3, 6, 6);
                    }
                }
            }
            else if (kind == "bar")
            {
                float groupH = plot.Height / categoryCount;
                float barH = Math.Max(2f, groupH * 0.75f / series.Count);

                for (int ci = 0; ci < categoryCount; ci++)
                {
                    for (int si = 0; si < series.Count; si++)
                    {
                        if (ci >= series[si].Values.Count) continue;

                        float y = plot.Top + ci * groupH + groupH * 0.12f + si * barH;
                        float w = (float)(plot.Width * Math.Abs(series[si].Values[ci]) / max);

                        using (Brush brush = new SolidBrush(palette[si % palette.Length]))
                            g.FillRectangle(brush, plot.Left, y, w, Math.Max(1f, barH - 1));
                    }
                }
            }
            else
            {
                float groupW = plot.Width / categoryCount;
                float barW = Math.Max(2f, groupW * 0.75f / series.Count);

                for (int ci = 0; ci < categoryCount; ci++)
                {
                    for (int si = 0; si < series.Count; si++)
                    {
                        if (ci >= series[si].Values.Count) continue;

                        float x = plot.Left + ci * groupW + groupW * 0.12f + si * barW;
                        float h = (float)(plot.Height * Math.Abs(series[si].Values[ci]) / max);

                        using (Brush brush = new SolidBrush(palette[si % palette.Length]))
                            g.FillRectangle(brush, x, plot.Bottom - h, Math.Max(1f, barW - 1), h);
                    }
                }
            }

            DrawChartCategoryLabels(g, plot, series[0], categoryCount, kind);
            DrawChartLegend(g, rect, series, palette);
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
            Color[] palette)
        {
            if (series.Count <= 1) return;

            float x = rect.Right - Math.Max(65f, rect.Width * 0.14f);
            float y = rect.Top + Math.Max(35f, rect.Height * 0.12f);

            using (Font font = SafeFont("Arial", Math.Max(7f, Math.Min(11f, rect.Height / 28f))))
            {
                for (int i = 0; i < series.Count; i++)
                {
                    using (Brush swatch = new SolidBrush(palette[i % palette.Length]))
                        g.FillRectangle(swatch, x, y + 3, 10, 10);

                    using (Brush text = new SolidBrush(Color.FromArgb(60, 60, 60)))
                        g.DrawString(
                            string.IsNullOrEmpty(series[i].Name) ? "Series " + (i + 1).ToString() : series[i].Name,
                            font,
                            text,
                            x + 14,
                            y);

                    y += font.Height + 4;
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
