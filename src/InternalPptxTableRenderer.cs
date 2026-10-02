using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Xml;

namespace PptxViewer
{
    internal static partial class InternalPptxRenderer
    {
        private static void DrawRichTable(
            Graphics g,
            XmlNode table,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            if (g == null || table == null || rect.Width <= 0f || rect.Height <= 0f)
                return;

            List<XmlNode> rows = new List<XmlNode>();
            foreach (XmlNode child in table.ChildNodes)
            {
                if (child.LocalName == "tr")
                    rows.Add(child);
            }

            if (rows.Count == 0)
                return;

            List<long> columnWidths = ReadTableColumnWidths(table);
            int inferredColumns = 0;

            for (int r = 0; r < rows.Count; r++)
            {
                int count = 0;
                foreach (XmlNode child in rows[r].ChildNodes)
                {
                    if (child.LocalName != "tc")
                        continue;

                    int span = Math.Max(1, (int)GetLong(child, "gridSpan", 1));
                    bool continuation = IsTableMergeContinuation(child);
                    count += continuation ? 1 : span;
                }
                inferredColumns = Math.Max(inferredColumns, count);
            }

            int columnCount = Math.Max(columnWidths.Count, inferredColumns);
            if (columnCount <= 0)
                return;

            while (columnWidths.Count < columnCount)
                columnWidths.Add(1);

            List<long> rowHeights = new List<long>();
            for (int r = 0; r < rows.Count; r++)
            {
                long height = GetLong(rows[r], "h", 0);
                rowHeights.Add(height > 0 ? height : 1);
            }

            float[] x = BuildTablePositions(rect.Left, rect.Width, columnWidths);
            float[] y = BuildTablePositions(rect.Top, rect.Height, rowHeights);

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                int columnIndex = 0;

                foreach (XmlNode cell in rows[rowIndex].ChildNodes)
                {
                    if (cell.LocalName != "tc")
                        continue;

                    bool continuation = IsTableMergeContinuation(cell);
                    int colSpan = Math.Max(1, (int)GetLong(cell, "gridSpan", 1));
                    int rowSpan = Math.Max(1, (int)GetLong(cell, "rowSpan", 1));

                    if (continuation)
                    {
                        columnIndex++;
                        continue;
                    }

                    int lastColumn = Math.Min(columnCount, columnIndex + colSpan);
                    int lastRow = Math.Min(rows.Count, rowIndex + rowSpan);

                    if (columnIndex >= columnCount || lastColumn <= columnIndex)
                        break;

                    RectangleF cellRect = new RectangleF(
                        x[columnIndex],
                        y[rowIndex],
                        Math.Max(1f, x[lastColumn] - x[columnIndex]),
                        Math.Max(1f, y[lastRow] - y[rowIndex]));

                    DrawRichTableCell(g, cell, cellRect, theme);
                    columnIndex += colSpan;
                }
            }
        }

        private static List<long> ReadTableColumnWidths(XmlNode table)
        {
            List<long> result = new List<long>();
            XmlNode grid = DirectChild(table, "tblGrid");
            if (grid == null)
                return result;

            foreach (XmlNode child in grid.ChildNodes)
            {
                if (child.LocalName != "gridCol")
                    continue;

                long width = GetLong(child, "w", 0);
                result.Add(width > 0 ? width : 1);
            }

            return result;
        }

        private static float[] BuildTablePositions(
            float start,
            float extent,
            List<long> sizes)
        {
            float[] positions = new float[sizes.Count + 1];
            positions[0] = start;

            double total = 0.0;
            for (int i = 0; i < sizes.Count; i++)
                total += Math.Max(1L, sizes[i]);

            if (total <= 0.0)
                total = sizes.Count;

            double cursor = start;
            for (int i = 0; i < sizes.Count; i++)
            {
                cursor += extent * Math.Max(1L, sizes[i]) / total;
                positions[i + 1] = (float)cursor;
            }

            positions[positions.Length - 1] = start + extent;
            return positions;
        }

        private static bool IsTableMergeContinuation(XmlNode cell)
        {
            string hMerge = GetAttr(cell, "hMerge");
            string vMerge = GetAttr(cell, "vMerge");
            return IsTableTrue(hMerge) || IsTableTrue(vMerge);
        }

        private static bool IsTableTrue(string value)
        {
            return value == "1" ||
                string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        private static void DrawRichTableCell(
            Graphics g,
            XmlNode cell,
            RectangleF rect,
            Dictionary<string, Color> theme)
        {
            XmlNode tcPr = DirectChild(cell, "tcPr");
            Color? fill = tcPr != null ? ReadSolidFill(tcPr, theme) : null;

            if (fill.HasValue)
            {
                using (Brush brush = new SolidBrush(fill.Value))
                    g.FillRectangle(brush, rect);
            }

            DrawTableBorder(g, tcPr, "lnT", rect.Left, rect.Top, rect.Right, rect.Top, theme);
            DrawTableBorder(g, tcPr, "lnB", rect.Left, rect.Bottom, rect.Right, rect.Bottom, theme);
            DrawTableBorder(g, tcPr, "lnL", rect.Left, rect.Top, rect.Left, rect.Bottom, theme);
            DrawTableBorder(g, tcPr, "lnR", rect.Right, rect.Top, rect.Right, rect.Bottom, theme);

            // If the cell does not carry explicit border lines, keep a very
            // light grid so plain tables remain readable instead of becoming
            // visually merged into one rectangle.
            if (tcPr == null ||
                (DirectChild(tcPr, "lnT") == null &&
                 DirectChild(tcPr, "lnB") == null &&
                 DirectChild(tcPr, "lnL") == null &&
                 DirectChild(tcPr, "lnR") == null))
            {
                using (Pen grid = new Pen(Color.FromArgb(90, 120, 120, 120), 1f))
                    g.DrawRectangle(grid, rect.X, rect.Y, rect.Width, rect.Height);
            }

            XmlNode txBody = DirectChild(cell, "txBody");
            if (txBody == null)
                return;

            XmlDocument owner = new XmlDocument();
            XmlElement pseudoShape = owner.CreateElement(
                "p",
                "sp",
                "http://schemas.openxmlformats.org/presentationml/2006/main");
            owner.AppendChild(pseudoShape);

            XmlNode textClone = owner.ImportNode(txBody, true);
            pseudoShape.AppendChild(textClone);

            XmlNode bodyPr = DirectChild(textClone, "bodyPr");
            if (bodyPr == null)
            {
                bodyPr = owner.CreateElement(
                    "a",
                    "bodyPr",
                    "http://schemas.openxmlformats.org/drawingml/2006/main");
                textClone.PrependChild(bodyPr);
            }

            long defaultMargin = 45720;
            SetTableTextBodyAttribute(bodyPr, "lIns", GetLong(tcPr, "marL", defaultMargin));
            SetTableTextBodyAttribute(bodyPr, "rIns", GetLong(tcPr, "marR", defaultMargin));
            SetTableTextBodyAttribute(bodyPr, "tIns", GetLong(tcPr, "marT", defaultMargin));
            SetTableTextBodyAttribute(bodyPr, "bIns", GetLong(tcPr, "marB", defaultMargin));

            string anchor = GetAttr(tcPr, "anchor");
            if (!string.IsNullOrEmpty(anchor))
                SetTableTextBodyAttribute(bodyPr, "anchor", anchor);

            string vertical = GetAttr(tcPr, "vert");
            if (!string.IsNullOrEmpty(vertical))
                SetTableTextBodyAttribute(bodyPr, "vert", vertical);

            DrawRichShapeText(g, pseudoShape, rect, theme, 1);
        }

        private static void SetTableTextBodyAttribute(
            XmlNode node,
            string name,
            long value)
        {
            SetTableTextBodyAttribute(node, name, value.ToString());
        }

        private static void SetTableTextBodyAttribute(
            XmlNode node,
            string name,
            string value)
        {
            if (node == null || node.OwnerDocument == null)
                return;

            XmlAttribute attribute = node.Attributes != null
                ? node.Attributes[name]
                : null;

            if (attribute == null)
            {
                attribute = node.OwnerDocument.CreateAttribute(name);
                node.Attributes.Append(attribute);
            }

            attribute.Value = value ?? "";
        }

        private static void DrawTableBorder(
            Graphics g,
            XmlNode tcPr,
            string lineName,
            float x1,
            float y1,
            float x2,
            float y2,
            Dictionary<string, Color> theme)
        {
            if (tcPr == null)
                return;

            XmlNode line = DirectChild(tcPr, lineName);
            if (line == null)
                return;

            if (DirectChild(line, "noFill") != null)
                return;

            Color color = ReadSolidFill(line, theme) ?? Color.FromArgb(120, 120, 120);
            long rawWidth = GetLong(line, "w", 12700);
            float width = Math.Max(1f, (float)(rawWidth / 12700.0 * 2.0));

            using (Pen pen = new Pen(color, width))
            {
                pen.LineJoin = LineJoin.Round;

                XmlNode dash = DirectChild(line, "prstDash");
                string dashValue = dash != null ? GetAttr(dash, "val") : "";

                if (dashValue == "dash" || dashValue == "lgDash")
                    pen.DashStyle = DashStyle.Dash;
                else if (dashValue == "dot" || dashValue == "sysDot")
                    pen.DashStyle = DashStyle.Dot;
                else if (dashValue == "dashDot" || dashValue == "sysDashDot")
                    pen.DashStyle = DashStyle.DashDot;

                g.DrawLine(pen, x1, y1, x2, y2);
            }
        }
    }
}
