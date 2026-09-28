using System;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxTableWriter
    {
        private const string PresentationNamespace =
            "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string DrawingNamespace =
            "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string TableGraphicDataUri =
            "http://schemas.openxmlformats.org/drawingml/2006/table";

        public static void InjectTables(
            PresentationDocument document,
            string pptxPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            if (string.IsNullOrEmpty(pptxPath))
                throw new ArgumentException("PPTX path is required.", "pptxPath");

            bool hasTables = false;

            for (int i = 0; i < document.Slides.Count; i++)
            {
                PresentationSlide slide = document.Slides[i];
                if (slide != null && slide.Tables.Count > 0)
                {
                    hasTables = true;
                    break;
                }
            }

            if (!hasTables)
                return;

            using (ZipArchive archive = ZipFile.Open(pptxPath, ZipArchiveMode.Update))
            {
                for (int slideIndex = 0;
                     slideIndex < document.Slides.Count;
                     slideIndex++)
                {
                    PresentationSlide slide = document.Slides[slideIndex];
                    if (slide == null || slide.Tables.Count == 0)
                        continue;

                    string partName =
                        "ppt/slides/slide" +
                        (slideIndex + 1).ToString() +
                        ".xml";

                    XmlDocument slideDocument =
                        OpcPackageUtility.ReadXmlPart(archive, partName);

                    if (slideDocument == null)
                    {
                        throw new InvalidOperationException(
                            "Cannot inject table because slide part is missing: " +
                            partName);
                    }

                    XmlNode shapeTree = FindFirst(
                        slideDocument.DocumentElement,
                        "spTree");

                    if (shapeTree == null)
                        throw new InvalidOperationException("Slide shape tree is missing.");

                    int nextShapeId = FindMaxShapeId(shapeTree) + 1;

                    for (int tableIndex = 0;
                         tableIndex < slide.Tables.Count;
                         tableIndex++)
                    {
                        PresentationTable table = slide.Tables[tableIndex];
                        if (table == null)
                            continue;

                        XmlElement frame = BuildGraphicFrame(
                            slideDocument,
                            table,
                            nextShapeId++);
                        shapeTree.AppendChild(frame);
                    }

                    OpcPackageUtility.WriteXmlPart(
                        archive,
                        partName,
                        slideDocument);
                }
            }
        }

        private static XmlElement BuildGraphicFrame(
            XmlDocument document,
            PresentationTable table,
            int shapeId)
        {
            XmlElement frame = P(document, "graphicFrame");

            XmlElement nv = P(document, "nvGraphicFramePr");
            XmlElement cNvPr = P(document, "cNvPr");
            cNvPr.SetAttribute("id", shapeId.ToString());
            cNvPr.SetAttribute(
                "name",
                string.IsNullOrEmpty(table.Name)
                    ? "Table " + shapeId.ToString()
                    : table.Name);
            nv.AppendChild(cNvPr);

            XmlElement cNvGraphic = P(document, "cNvGraphicFramePr");
            XmlElement locks = A(document, "graphicFrameLocks");
            locks.SetAttribute("noGrp", "1");
            cNvGraphic.AppendChild(locks);
            nv.AppendChild(cNvGraphic);
            nv.AppendChild(P(document, "nvPr"));
            frame.AppendChild(nv);

            XmlElement transform = P(document, "xfrm");
            XmlElement offset = A(document, "off");
            offset.SetAttribute("x", Math.Max(0L, table.X).ToString());
            offset.SetAttribute("y", Math.Max(0L, table.Y).ToString());
            transform.AppendChild(offset);

            XmlElement extent = A(document, "ext");
            extent.SetAttribute("cx", Math.Max(1L, table.Width).ToString());
            extent.SetAttribute("cy", Math.Max(1L, table.Height).ToString());
            transform.AppendChild(extent);
            frame.AppendChild(transform);

            XmlElement graphic = A(document, "graphic");
            XmlElement graphicData = A(document, "graphicData");
            graphicData.SetAttribute("uri", TableGraphicDataUri);
            graphicData.AppendChild(BuildTable(document, table));
            graphic.AppendChild(graphicData);
            frame.AppendChild(graphic);

            return frame;
        }

        private static XmlElement BuildTable(
            XmlDocument document,
            PresentationTable table)
        {
            int rows = Math.Max(1, table.Rows);
            int columns = Math.Max(1, table.Columns);
            long totalWidth = Math.Max(1L, table.Width);
            long totalHeight = Math.Max(1L, table.Height);
            long columnWidth = Math.Max(1L, totalWidth / columns);
            long rowHeight = Math.Max(1L, totalHeight / rows);

            XmlElement tbl = A(document, "tbl");
            XmlElement properties = A(document, "tblPr");
            properties.SetAttribute("firstRow", "1");
            properties.SetAttribute("bandRow", "1");
            tbl.AppendChild(properties);

            XmlElement grid = A(document, "tblGrid");
            long assignedWidth = 0L;

            for (int column = 0; column < columns; column++)
            {
                long width = column == columns - 1
                    ? Math.Max(1L, totalWidth - assignedWidth)
                    : columnWidth;

                XmlElement gridColumn = A(document, "gridCol");
                gridColumn.SetAttribute("w", width.ToString());
                grid.AppendChild(gridColumn);
                assignedWidth += width;
            }

            tbl.AppendChild(grid);
            long assignedHeight = 0L;

            for (int row = 0; row < rows; row++)
            {
                long height = row == rows - 1
                    ? Math.Max(1L, totalHeight - assignedHeight)
                    : rowHeight;

                XmlElement tr = A(document, "tr");
                tr.SetAttribute("h", height.ToString());

                for (int column = 0; column < columns; column++)
                {
                    PresentationTableCell cell =
                        table.GetCell(row, column) ??
                        new PresentationTableCell();
                    tr.AppendChild(BuildCell(document, cell));
                }

                tbl.AppendChild(tr);
                assignedHeight += height;
            }

            return tbl;
        }

        private static XmlElement BuildCell(
            XmlDocument document,
            PresentationTableCell cell)
        {
            XmlElement tc = A(document, "tc");
            XmlElement txBody = A(document, "txBody");
            txBody.AppendChild(A(document, "bodyPr"));
            txBody.AppendChild(A(document, "lstStyle"));

            string text = (cell.Text ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');
            string[] lines = text.Split('\n');

            if (lines.Length == 0)
                lines = new string[] { string.Empty };

            for (int i = 0; i < lines.Length; i++)
            {
                XmlElement paragraph = A(document, "p");
                XmlElement pPr = A(document, "pPr");
                pPr.SetAttribute("algn", AlignmentValue(cell.Alignment));
                paragraph.AppendChild(pPr);

                if (!string.IsNullOrEmpty(lines[i]))
                {
                    XmlElement run = A(document, "r");
                    XmlElement rPr = A(document, "rPr");
                    rPr.SetAttribute("lang", "ko-KR");
                    rPr.SetAttribute(
                        "sz",
                        Math.Max(
                            100,
                            Math.Min(
                                40000,
                                (int)Math.Round(cell.FontSizePoints * 100.0f)))
                            .ToString());

                    if (cell.Bold)
                        rPr.SetAttribute("b", "1");

                    XmlElement fill = A(document, "solidFill");
                    XmlElement color = A(document, "srgbClr");
                    color.SetAttribute("val", NormalizeColor(cell.TextColorHex, "20242A"));
                    fill.AppendChild(color);
                    rPr.AppendChild(fill);

                    string family = string.IsNullOrEmpty(cell.FontFamily)
                        ? "Arial"
                        : cell.FontFamily;
                    XmlElement latin = A(document, "latin");
                    latin.SetAttribute("typeface", family);
                    rPr.AppendChild(latin);
                    XmlElement eastAsian = A(document, "ea");
                    eastAsian.SetAttribute("typeface", family);
                    rPr.AppendChild(eastAsian);
                    run.AppendChild(rPr);

                    XmlElement textNode = A(document, "t");
                    textNode.AppendChild(document.CreateTextNode(lines[i]));
                    run.AppendChild(textNode);
                    paragraph.AppendChild(run);
                }

                XmlElement end = A(document, "endParaRPr");
                end.SetAttribute("lang", "ko-KR");
                end.SetAttribute(
                    "sz",
                    Math.Max(
                        100,
                        Math.Min(
                            40000,
                            (int)Math.Round(cell.FontSizePoints * 100.0f)))
                        .ToString());
                paragraph.AppendChild(end);
                txBody.AppendChild(paragraph);
            }

            tc.AppendChild(txBody);

            XmlElement tcPr = A(document, "tcPr");
            XmlElement cellFill = A(document, "solidFill");
            XmlElement cellColor = A(document, "srgbClr");
            cellColor.SetAttribute("val", NormalizeColor(cell.FillColorHex, "FFFFFF"));
            cellFill.AppendChild(cellColor);
            tcPr.AppendChild(cellFill);
            tc.AppendChild(tcPr);

            return tc;
        }

        private static int FindMaxShapeId(XmlNode node)
        {
            if (node == null)
                return 1;

            int max = 1;
            FindMaxShapeIdRecursive(node, ref max);
            return max;
        }

        private static void FindMaxShapeIdRecursive(
            XmlNode node,
            ref int max)
        {
            if (node == null)
                return;

            if (node.LocalName == "cNvPr")
            {
                int id;
                XmlAttribute attribute =
                    node.Attributes == null ? null : node.Attributes["id"];

                if (attribute != null &&
                    int.TryParse(attribute.Value, out id) &&
                    id > max)
                {
                    max = id;
                }
            }

            for (int i = 0; i < node.ChildNodes.Count; i++)
                FindMaxShapeIdRecursive(node.ChildNodes[i], ref max);
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

        private static XmlElement P(
            XmlDocument document,
            string localName)
        {
            return document.CreateElement(
                "p",
                localName,
                PresentationNamespace);
        }

        private static XmlElement A(
            XmlDocument document,
            string localName)
        {
            return document.CreateElement(
                "a",
                localName,
                DrawingNamespace);
        }

        private static string AlignmentValue(
            PresentationTextAlignment alignment)
        {
            if (alignment == PresentationTextAlignment.Center)
                return "ctr";
            if (alignment == PresentationTextAlignment.Right)
                return "r";
            return "l";
        }

        private static string NormalizeColor(
            string value,
            string fallback)
        {
            string candidate = (value ?? string.Empty)
                .Trim()
                .TrimStart('#')
                .ToUpperInvariant();

            if (candidate.Length != 6)
                return fallback;

            for (int i = 0; i < candidate.Length; i++)
            {
                char ch = candidate[i];
                bool isHex =
                    (ch >= '0' && ch <= '9') ||
                    (ch >= 'A' && ch <= 'F');

                if (!isHex)
                    return fallback;
            }

            return candidate;
        }
    }
}
