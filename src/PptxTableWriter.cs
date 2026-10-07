using System;
using System.Collections.Generic;
using System.IO;
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
        private const string ChartGraphicDataUri =
            "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string ChartNamespace =
            "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string RelationshipsNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string ChartRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        private const string ChartContentType =
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        public static void InjectTables(
            PresentationDocument document,
            string pptxPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            if (string.IsNullOrEmpty(pptxPath))
                throw new ArgumentException("PPTX path is required.", "pptxPath");

            bool hasGraphicFrames = false;
            bool hasCharts = false;

            for (int i = 0; i < document.Slides.Count; i++)
            {
                PresentationSlide slide = document.Slides[i];

                if (slide == null)
                    continue;

                if (slide.Tables.Count > 0 ||
                    slide.Charts.Count > 0)
                {
                    hasGraphicFrames = true;
                }

                if (slide.Charts.Count > 0)
                    hasCharts = true;
            }

            if (!hasGraphicFrames)
                return;

            using (ZipArchive archive = ZipFile.Open(pptxPath, ZipArchiveMode.Update))
            {
                for (int slideIndex = 0;
                     slideIndex < document.Slides.Count;
                     slideIndex++)
                {
                    PresentationSlide slide = document.Slides[slideIndex];
                    if (slide == null ||
                        (slide.Tables.Count == 0 &&
                         slide.Charts.Count == 0))
                    {
                        continue;
                    }

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
                    XmlElement[] tableFrames =
                        new XmlElement[slide.Tables.Count];
                    XmlElement[] chartFrames =
                        new XmlElement[slide.Charts.Count];

                    for (int tableIndex = 0;
                         tableIndex < slide.Tables.Count;
                         tableIndex++)
                    {
                        PresentationTable table = slide.Tables[tableIndex];
                        if (table == null)
                            continue;

                        tableFrames[tableIndex] =
                            BuildGraphicFrame(
                                slideDocument,
                                table,
                                nextShapeId++);
                    }

                    for (int chartIndex = 0;
                         chartIndex < slide.Charts.Count;
                         chartIndex++)
                    {
                        PresentationChart chartModel =
                            slide.Charts[chartIndex];

                        if (chartModel == null ||
                            !PptxChartReader.IsSelfContainedChartXml(
                                chartModel.XmlData))
                        {
                            throw new InvalidOperationException(
                                "Chart payload is missing or requires related parts that are not yet editable.");
                        }

                        chartFrames[chartIndex] =
                            BuildChartGraphicFrame(
                                slideDocument,
                                chartModel,
                                nextShapeId++,
                                ChartRelationshipId(
                                    chartIndex));
                    }

                    // PptxWriter already emits text/shape/image nodes in
                    // ObjectOrder while skipping graphic-frame kinds.
                    // Rebuild the object section after table/chart XML is
                    // available so every editable object shares one z-order.
                    List<XmlNode> generatedObjects =
                        new List<XmlNode>();

                    for (int childIndex = 0;
                         childIndex < shapeTree.ChildNodes.Count;
                         childIndex++)
                    {
                        XmlNode child =
                            shapeTree.ChildNodes[childIndex];

                        if (child == null)
                            continue;

                        if (child.LocalName == "sp" ||
                            child.LocalName == "pic")
                        {
                            generatedObjects.Add(child);
                        }
                    }

                    for (int i = 0;
                         i < generatedObjects.Count;
                         i++)
                    {
                        shapeTree.RemoveChild(
                            generatedObjects[i]);
                    }

                    slide.SynchronizeObjectOrder();
                    int generatedIndex = 0;

                    for (int orderIndex = 0;
                         orderIndex < slide.ObjectOrder.Count;
                         orderIndex++)
                    {
                        PresentationLayerEntry entry =
                            slide.ObjectOrder[orderIndex];

                        if (entry == null)
                            continue;

                        if (entry.Kind ==
                                PresentationLayerKind.Table &&
                            entry.Index >= 0 &&
                            entry.Index < tableFrames.Length &&
                            tableFrames[entry.Index] != null)
                        {
                            shapeTree.AppendChild(
                                tableFrames[entry.Index]);
                        }
                        else if (entry.Kind ==
                                     PresentationLayerKind.Chart &&
                                 entry.Index >= 0 &&
                                 entry.Index < chartFrames.Length &&
                                 chartFrames[entry.Index] != null)
                        {
                            shapeTree.AppendChild(
                                chartFrames[entry.Index]);
                        }
                        else if (entry.Kind !=
                                     PresentationLayerKind.Table &&
                                 entry.Kind !=
                                     PresentationLayerKind.Chart &&
                                 generatedIndex <
                                     generatedObjects.Count)
                        {
                            shapeTree.AppendChild(
                                generatedObjects[
                                    generatedIndex++]);
                        }
                    }

                    while (generatedIndex <
                           generatedObjects.Count)
                    {
                        shapeTree.AppendChild(
                            generatedObjects[
                                generatedIndex++]);
                    }

                    OpcPackageUtility.WriteXmlPart(
                        archive,
                        partName,
                        slideDocument);

                    if (slide.Charts.Count > 0)
                    {
                        WriteChartsForSlide(
                            archive,
                            slideIndex,
                            slide);
                    }
                }

                if (hasCharts)
                {
                    EnsureChartContentTypes(
                        archive,
                        document);
                }
            }
        }

        private static void WriteChartsForSlide(
            ZipArchive archive,
            int slideIndex,
            PresentationSlide slide)
        {
            string sourcePart =
                "ppt/slides/slide" +
                (slideIndex + 1).ToString() +
                ".xml";
            List<OpcRelationship> relationships =
                OpcPackageUtility.ReadRelationships(
                    archive,
                    sourcePart);

            for (int chartIndex = 0;
                 chartIndex < slide.Charts.Count;
                 chartIndex++)
            {
                PresentationChart chart =
                    slide.Charts[chartIndex];

                if (chart == null ||
                    !PptxChartReader.IsSelfContainedChartXml(
                        chart.XmlData))
                {
                    throw new InvalidOperationException(
                        "Cannot write a chart that is not self-contained.");
                }

                string chartPart =
                    GetChartPartName(
                        slideIndex,
                        chartIndex);

                WriteRawPart(
                    archive,
                    chartPart,
                    chart.XmlData);

                OpcRelationship relationship =
                    new OpcRelationship();
                relationship.Id =
                    ChartRelationshipId(
                        chartIndex);
                relationship.Type =
                    ChartRelationship;
                relationship.Target =
                    "../charts/" +
                    Path.GetFileName(
                        chartPart);
                relationships.Add(
                    relationship);
            }

            XmlDocument relationshipDocument =
                OpcPackageUtility
                    .CreateRelationshipsDocument(
                        relationships);

            OpcPackageUtility.WriteXmlPart(
                archive,
                OpcPackageUtility
                    .GetRelationshipPartName(
                        sourcePart),
                relationshipDocument);
        }

        private static void EnsureChartContentTypes(
            ZipArchive archive,
            PresentationDocument document)
        {
            XmlDocument contentTypes =
                OpcPackageUtility.ReadXmlPart(
                    archive,
                    "[Content_Types].xml");

            if (contentTypes == null ||
                contentTypes.DocumentElement == null)
            {
                throw new InvalidOperationException(
                    "PPTX content types part is missing.");
            }

            XmlElement root =
                contentTypes.DocumentElement;
            string ns =
                root.NamespaceURI;

            for (int slideIndex = 0;
                 slideIndex < document.Slides.Count;
                 slideIndex++)
            {
                PresentationSlide slide =
                    document.Slides[slideIndex];

                if (slide == null)
                    continue;

                for (int chartIndex = 0;
                     chartIndex < slide.Charts.Count;
                     chartIndex++)
                {
                    string partName =
                        "/" +
                        GetChartPartName(
                            slideIndex,
                            chartIndex);

                    if (HasContentTypeOverride(
                            root,
                            partName))
                    {
                        continue;
                    }

                    XmlElement element =
                        contentTypes.CreateElement(
                            "Override",
                            ns);
                    element.SetAttribute(
                        "PartName",
                        partName);
                    element.SetAttribute(
                        "ContentType",
                        ChartContentType);
                    root.AppendChild(element);
                }
            }

            OpcPackageUtility.WriteXmlPart(
                archive,
                "[Content_Types].xml",
                contentTypes);
        }

        private static bool HasContentTypeOverride(
            XmlElement root,
            string partName)
        {
            if (root == null)
                return false;

            for (int i = 0;
                 i < root.ChildNodes.Count;
                 i++)
            {
                XmlNode node =
                    root.ChildNodes[i];

                if (node.LocalName !=
                    "Override" ||
                    node.Attributes == null)
                {
                    continue;
                }

                XmlAttribute attribute =
                    node.Attributes[
                        "PartName"];

                if (attribute != null &&
                    string.Equals(
                        attribute.Value,
                        partName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static void WriteRawPart(
            ZipArchive archive,
            string partName,
            byte[] data)
        {
            string normalized =
                OpcPackageUtility
                    .NormalizePartName(
                        partName);
            ZipArchiveEntry existing =
                archive.GetEntry(
                    normalized);

            if (existing != null)
                existing.Delete();

            ZipArchiveEntry entry =
                archive.CreateEntry(
                    normalized,
                    CompressionLevel.Optimal);

            using (Stream stream =
                entry.Open())
            {
                stream.Write(
                    data,
                    0,
                    data.Length);
            }
        }

        private static string GetChartPartName(
            int slideIndex,
            int chartIndex)
        {
            return
                "ppt/charts/slide" +
                (slideIndex + 1).ToString() +
                "_chart" +
                (chartIndex + 1).ToString() +
                ".xml";
        }

        private static string ChartRelationshipId(
            int chartIndex)
        {
            return
                "rIdChart" +
                (chartIndex + 1).ToString();
        }

        private static XmlElement BuildChartGraphicFrame(
            XmlDocument document,
            PresentationChart chart,
            int shapeId,
            string relationshipId)
        {
            XmlElement frame =
                P(
                    document,
                    "graphicFrame");

            XmlElement nv =
                P(
                    document,
                    "nvGraphicFramePr");
            XmlElement cNvPr =
                P(
                    document,
                    "cNvPr");
            cNvPr.SetAttribute(
                "id",
                shapeId.ToString());
            cNvPr.SetAttribute(
                "name",
                string.IsNullOrEmpty(
                    chart.Name)
                    ? "Chart " +
                      shapeId.ToString()
                    : chart.Name);
            nv.AppendChild(cNvPr);

            XmlElement cNvGraphic =
                P(
                    document,
                    "cNvGraphicFramePr");
            XmlElement locks =
                A(
                    document,
                    "graphicFrameLocks");
            locks.SetAttribute(
                "noGrp",
                "1");
            cNvGraphic.AppendChild(
                locks);
            nv.AppendChild(
                cNvGraphic);
            nv.AppendChild(
                P(
                    document,
                    "nvPr"));
            frame.AppendChild(nv);

            XmlElement transform =
                P(
                    document,
                    "xfrm");
            XmlElement offset =
                A(
                    document,
                    "off");
            offset.SetAttribute(
                "x",
                Math.Max(
                    0L,
                    chart.X).ToString());
            offset.SetAttribute(
                "y",
                Math.Max(
                    0L,
                    chart.Y).ToString());
            transform.AppendChild(
                offset);

            XmlElement extent =
                A(
                    document,
                    "ext");
            extent.SetAttribute(
                "cx",
                Math.Max(
                    1L,
                    chart.Width).ToString());
            extent.SetAttribute(
                "cy",
                Math.Max(
                    1L,
                    chart.Height).ToString());
            transform.AppendChild(
                extent);
            frame.AppendChild(
                transform);

            XmlElement graphic =
                A(
                    document,
                    "graphic");
            XmlElement graphicData =
                A(
                    document,
                    "graphicData");
            graphicData.SetAttribute(
                "uri",
                ChartGraphicDataUri);

            XmlElement chartReference =
                document.CreateElement(
                    "c",
                    "chart",
                    ChartNamespace);
            XmlAttribute relationshipAttribute =
                document.CreateAttribute(
                    "r",
                    "id",
                    RelationshipsNamespace);
            relationshipAttribute.Value =
                relationshipId;
            chartReference.Attributes.Append(
                relationshipAttribute);
            graphicData.AppendChild(
                chartReference);
            graphic.AppendChild(
                graphicData);
            frame.AppendChild(
                graphic);

            return frame;
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
