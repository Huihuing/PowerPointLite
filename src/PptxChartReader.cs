using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxChartReader
    {
        private const string ChartRelationshipSuffix =
            "/chart";
        private const string ChartGraphicDataUri =
            "http://schemas.openxmlformats.org/drawingml/2006/chart";

        public static PresentationChart Read(
            ZipArchive archive,
            XmlNode graphicFrame,
            Dictionary<string, OpcRelationship> relationships)
        {
            if (archive == null ||
                graphicFrame == null ||
                relationships == null)
            {
                return null;
            }

            XmlNode graphicData =
                FindFirst(
                    graphicFrame,
                    "graphicData");
            string uri =
                GetAttribute(
                    graphicData,
                    "uri");

            if (!string.IsNullOrEmpty(uri) &&
                !string.Equals(
                    uri,
                    ChartGraphicDataUri,
                    StringComparison.Ordinal))
            {
                return null;
            }

            XmlNode chartReference =
                FindFirst(
                    graphicData,
                    "chart");
            string relationshipId =
                GetRelationshipReference(
                    chartReference,
                    "id");
            OpcRelationship relationship;

            if (string.IsNullOrEmpty(relationshipId) ||
                !relationships.TryGetValue(
                    relationshipId,
                    out relationship) ||
                relationship == null ||
                relationship.IsExternal ||
                !relationship.Type.EndsWith(
                    ChartRelationshipSuffix,
                    StringComparison.Ordinal))
            {
                return null;
            }

            // First editable milestone intentionally accepts only chart
            // parts that are self-contained. Embedded workbooks,
            // externalData, userShapes, chart styles/colors and any other
            // related parts stay deny-by-default until their package graph
            // can be preserved losslessly.
            List<OpcRelationship> chartRelationships =
                OpcPackageUtility.ReadRelationships(
                    archive,
                    relationship.ResolvedPart);

            if (chartRelationships.Count != 0)
                return null;

            ZipArchiveEntry entry =
                archive.GetEntry(
                    OpcPackageUtility.NormalizePartName(
                        relationship.ResolvedPart));

            if (entry == null)
                return null;

            byte[] data;

            using (Stream input = entry.Open())
            using (MemoryStream output =
                new MemoryStream())
            {
                input.CopyTo(output);
                data = output.ToArray();
            }

            if (!IsSelfContainedChartXml(data))
                return null;

            PresentationChart chart =
                new PresentationChart();
            chart.XmlData = data;

            XmlNode nonVisual =
                FindFirst(
                    graphicFrame,
                    "cNvPr");
            string name =
                GetAttribute(
                    nonVisual,
                    "name");

            if (!string.IsNullOrEmpty(name))
                chart.Name = name;

            XmlNode transform =
                FindFirst(
                    graphicFrame,
                    "xfrm");
            XmlNode offset =
                FindDirectChild(
                    transform,
                    "off");
            XmlNode extent =
                FindDirectChild(
                    transform,
                    "ext");

            chart.X =
                GetLongAttribute(
                    offset,
                    "x",
                    chart.X);
            chart.Y =
                GetLongAttribute(
                    offset,
                    "y",
                    chart.Y);
            chart.Width =
                Math.Max(
                    1L,
                    GetLongAttribute(
                        extent,
                        "cx",
                        chart.Width));
            chart.Height =
                Math.Max(
                    1L,
                    GetLongAttribute(
                        extent,
                        "cy",
                        chart.Height));

            return chart;
        }

        private static bool IsSelfContainedChartXml(
            byte[] data)
        {
            if (data == null ||
                data.Length == 0)
            {
                return false;
            }

            try
            {
                XmlDocument document =
                    new XmlDocument();
                document.PreserveWhitespace = true;

                using (MemoryStream stream =
                    new MemoryStream(
                        data,
                        false))
                {
                    document.Load(stream);
                }

                XmlNode root =
                    document.DocumentElement;

                if (root == null ||
                    root.LocalName !=
                        "chartSpace")
                {
                    return false;
                }

                if (FindFirst(
                        root,
                        "externalData") != null ||
                    FindFirst(
                        root,
                        "pivotSource") != null ||
                    FindFirst(
                        root,
                        "userShapes") != null)
                {
                    return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetRelationshipReference(
            XmlNode node,
            string localName)
        {
            if (node == null ||
                node.Attributes == null)
            {
                return string.Empty;
            }

            for (int i = 0;
                 i < node.Attributes.Count;
                 i++)
            {
                XmlAttribute attribute =
                    node.Attributes[i];

                if (attribute.LocalName ==
                        localName &&
                    (attribute.Prefix == "r" ||
                     attribute.NamespaceURI.IndexOf(
                         "relationships",
                         StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return attribute.Value;
                }
            }

            return string.Empty;
        }

        private static XmlNode FindDirectChild(
            XmlNode node,
            string localName)
        {
            if (node == null)
                return null;

            for (int i = 0;
                 i < node.ChildNodes.Count;
                 i++)
            {
                XmlNode child =
                    node.ChildNodes[i];

                if (child.LocalName ==
                    localName)
                {
                    return child;
                }
            }

            return null;
        }

        private static XmlNode FindFirst(
            XmlNode node,
            string localName)
        {
            if (node == null)
                return null;

            if (node.LocalName ==
                localName)
            {
                return node;
            }

            for (int i = 0;
                 i < node.ChildNodes.Count;
                 i++)
            {
                XmlNode found =
                    FindFirst(
                        node.ChildNodes[i],
                        localName);

                if (found != null)
                    return found;
            }

            return null;
        }

        private static string GetAttribute(
            XmlNode node,
            string name)
        {
            if (node == null ||
                node.Attributes == null)
            {
                return string.Empty;
            }

            XmlAttribute direct =
                node.Attributes[name];

            if (direct != null)
                return direct.Value;

            for (int i = 0;
                 i < node.Attributes.Count;
                 i++)
            {
                XmlAttribute attribute =
                    node.Attributes[i];

                if (attribute.LocalName ==
                    name)
                {
                    return attribute.Value;
                }
            }

            return string.Empty;
        }

        private static long GetLongAttribute(
            XmlNode node,
            string name,
            long fallback)
        {
            long value;

            return long.TryParse(
                GetAttribute(
                    node,
                    name),
                out value)
                ? value
                : fallback;
        }
    }
}
