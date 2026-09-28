using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxTableReader
    {
        private const string TableGraphicDataUri =
            "http://schemas.openxmlformats.org/drawingml/2006/table";

        public static PresentationTable Read(XmlNode graphicFrame)
        {
            if (graphicFrame == null || graphicFrame.LocalName != "graphicFrame")
                return null;

            XmlNode graphicData = FindFirst(graphicFrame, "graphicData");
            string uri = GetAttribute(graphicData, "uri");
            XmlNode tableNode = FindFirst(graphicData, "tbl");

            if (tableNode == null ||
                (!string.IsNullOrEmpty(uri) &&
                 !string.Equals(
                    uri,
                    TableGraphicDataUri,
                    StringComparison.Ordinal)))
            {
                return null;
            }

            XmlNode grid = FindFirst(tableNode, "tblGrid");
            int columns = CountDirectChildren(grid, "gridCol");
            List<XmlNode> rows = DirectChildren(tableNode, "tr");

            if (columns <= 0 && rows.Count > 0)
                columns = CountDirectChildren(rows[0], "tc");

            if (rows.Count <= 0 || columns <= 0)
                return null;

            PresentationTable table =
                new PresentationTable(rows.Count, columns);

            XmlNode nonVisual = FindFirst(graphicFrame, "cNvPr");
            string name = GetAttribute(nonVisual, "name");
            if (!string.IsNullOrEmpty(name))
                table.Name = name;

            XmlNode transform = FindFirst(graphicFrame, "xfrm");
            XmlNode offset = FindDirectChild(transform, "off");
            XmlNode extent = FindDirectChild(transform, "ext");

            table.X = GetLongAttribute(offset, "x", table.X);
            table.Y = GetLongAttribute(offset, "y", table.Y);
            table.Width = Math.Max(
                1L,
                GetLongAttribute(extent, "cx", table.Width));
            table.Height = Math.Max(
                1L,
                GetLongAttribute(extent, "cy", table.Height));

            for (int row = 0; row < rows.Count; row++)
            {
                List<XmlNode> cells = DirectChildren(rows[row], "tc");

                for (int column = 0;
                     column < columns && column < cells.Count;
                     column++)
                {
                    PresentationTableCell target =
                        table.GetCell(row, column);

                    if (target != null)
                        ReadCell(cells[column], target);
                }
            }

            return table;
        }

        private static void ReadCell(
            XmlNode cellNode,
            PresentationTableCell target)
        {
            XmlNode textBody = FindDirectChild(cellNode, "txBody");
            target.Text = ReadText(textBody);

            XmlNode paragraphProperties = FindFirst(textBody, "pPr");
            string alignment = GetAttribute(paragraphProperties, "algn");

            if (alignment == "ctr")
                target.Alignment = PresentationTextAlignment.Center;
            else if (alignment == "r")
                target.Alignment = PresentationTextAlignment.Right;
            else
                target.Alignment = PresentationTextAlignment.Left;

            XmlNode runProperties = FindFirst(textBody, "rPr");
            if (runProperties == null)
                runProperties = FindFirst(textBody, "endParaRPr");

            int sizeHundredths =
                GetIntAttribute(runProperties, "sz", 1600);
            target.FontSizePoints = Math.Max(
                1f,
                sizeHundredths / 100f);
            target.Bold = IsTrue(GetAttribute(runProperties, "b"));

            XmlNode latin = FindFirst(runProperties, "latin");
            string typeface = GetAttribute(latin, "typeface");
            if (!string.IsNullOrEmpty(typeface))
                target.FontFamily = typeface;

            XmlNode textFill = FindDirectChild(runProperties, "solidFill");
            XmlNode textColor = FindFirst(textFill, "srgbClr");
            string textColorValue = GetAttribute(textColor, "val");
            if (!string.IsNullOrEmpty(textColorValue))
                target.TextColorHex = textColorValue;

            XmlNode cellProperties = FindDirectChild(cellNode, "tcPr");
            XmlNode cellFill = FindDirectChild(cellProperties, "solidFill");
            XmlNode fillColor = FindFirst(cellFill, "srgbClr");
            string fillValue = GetAttribute(fillColor, "val");
            if (!string.IsNullOrEmpty(fillValue))
                target.FillColorHex = fillValue;
        }

        private static string ReadText(XmlNode textBody)
        {
            if (textBody == null)
                return string.Empty;

            List<XmlNode> paragraphs = DirectChildren(textBody, "p");
            StringBuilder result = new StringBuilder();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (i > 0)
                    result.AppendLine();

                List<XmlNode> textNodes = FindAll(paragraphs[i], "t");

                for (int t = 0; t < textNodes.Count; t++)
                    result.Append(textNodes[t].InnerText);
            }

            return result.ToString();
        }

        private static int CountDirectChildren(
            XmlNode node,
            string localName)
        {
            return DirectChildren(node, localName).Count;
        }

        private static List<XmlNode> DirectChildren(
            XmlNode node,
            string localName)
        {
            List<XmlNode> result = new List<XmlNode>();

            if (node == null)
                return result;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                if (node.ChildNodes[i].LocalName == localName)
                    result.Add(node.ChildNodes[i]);
            }

            return result;
        }

        private static XmlNode FindDirectChild(
            XmlNode node,
            string localName)
        {
            if (node == null)
                return null;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                if (node.ChildNodes[i].LocalName == localName)
                    return node.ChildNodes[i];
            }

            return null;
        }

        private static XmlNode FindFirst(
            XmlNode node,
            string localName)
        {
            if (node == null)
                return null;

            if (node.LocalName == localName)
                return node;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                XmlNode found = FindFirst(node.ChildNodes[i], localName);
                if (found != null)
                    return found;
            }

            return null;
        }

        private static List<XmlNode> FindAll(
            XmlNode node,
            string localName)
        {
            List<XmlNode> result = new List<XmlNode>();
            AddMatches(node, localName, result);
            return result;
        }

        private static void AddMatches(
            XmlNode node,
            string localName,
            List<XmlNode> result)
        {
            if (node == null)
                return;

            if (node.LocalName == localName)
                result.Add(node);

            for (int i = 0; i < node.ChildNodes.Count; i++)
                AddMatches(node.ChildNodes[i], localName, result);
        }

        private static string GetAttribute(
            XmlNode node,
            string name)
        {
            if (node == null || node.Attributes == null)
                return string.Empty;

            XmlAttribute attribute = node.Attributes[name];
            if (attribute != null)
                return attribute.Value;

            for (int i = 0; i < node.Attributes.Count; i++)
            {
                if (node.Attributes[i].LocalName == name)
                    return node.Attributes[i].Value;
            }

            return string.Empty;
        }

        private static long GetLongAttribute(
            XmlNode node,
            string name,
            long fallback)
        {
            long value;
            return long.TryParse(GetAttribute(node, name), out value)
                ? value
                : fallback;
        }

        private static int GetIntAttribute(
            XmlNode node,
            string name,
            int fallback)
        {
            int value;
            return int.TryParse(GetAttribute(node, name), out value)
                ? value
                : fallback;
        }

        private static bool IsTrue(string value)
        {
            return value == "1" ||
                string.Equals(
                    value,
                    "true",
                    StringComparison.OrdinalIgnoreCase);
        }
    }
}
