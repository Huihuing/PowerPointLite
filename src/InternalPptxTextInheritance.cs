using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal static partial class InternalPptxRenderer
    {
        private sealed class PlaceholderTextStyleContext
        {
            public readonly List<XmlNode>[] Levels = new List<XmlNode>[9];
            public readonly List<XmlNode> BodyProperties = new List<XmlNode>();

            public PlaceholderTextStyleContext()
            {
                for (int i = 0; i < Levels.Length; i++)
                    Levels[i] = new List<XmlNode>();
            }
        }

        [ThreadStatic]
        private static Dictionary<string, PlaceholderTextStyleContext> activePlaceholderTextStyles;

        [ThreadStatic]
        private static PlaceholderTextStyleContext activePresentationDefaultTextStyle;

        [ThreadStatic]
        private static PlaceholderTextStyleContext activeMasterTitleStyle;

        [ThreadStatic]
        private static PlaceholderTextStyleContext activeMasterBodyStyle;

        [ThreadStatic]
        private static PlaceholderTextStyleContext activeMasterOtherStyle;

        private static void PrepareRichTextInheritance(
            ZipArchive zip,
            List<string> layers)
        {
            activePlaceholderTextStyles =
                new Dictionary<string, PlaceholderTextStyleContext>(StringComparer.OrdinalIgnoreCase);
            activePresentationDefaultTextStyle =
                new PlaceholderTextStyleContext();
            activeMasterTitleStyle =
                new PlaceholderTextStyleContext();
            activeMasterBodyStyle =
                new PlaceholderTextStyleContext();
            activeMasterOtherStyle =
                new PlaceholderTextStyleContext();

            if (zip == null || layers == null)
                return;

            XmlDocument presentation =
                LoadXml(
                    zip,
                    "ppt/presentation.xml");

            ReadPresentationDefaultTextStyle(
                presentation);

            // Master/layout are processed before the slide. Sources are kept
            // in that order so a more specific layout can override master
            // defaults without copying proprietary template assets.
            for (int layerIndex = 0; layerIndex < layers.Count - 1; layerIndex++)
            {
                string part = layers[layerIndex];
                XmlDocument document = LoadXml(zip, part);
                if (document == null)
                    continue;

                if (part.IndexOf("/slideMasters/", StringComparison.OrdinalIgnoreCase) >= 0)
                    ReadMasterTextStyles(document);

                XmlNode spTree = FindFirst(document, "spTree");
                if (spTree == null)
                    continue;

                foreach (XmlNode shape in FindAll(spTree, "sp"))
                {
                    XmlNode placeholder = FindFirst(shape, "ph");
                    if (placeholder == null)
                        continue;

                    string idx = GetAttr(placeholder, "idx");
                    string type = GetAttr(placeholder, "type");
                    XmlNode txBody = DirectChild(shape, "txBody");
                    XmlNode listStyle = txBody != null ? DirectChild(txBody, "lstStyle") : null;

                    if (!string.IsNullOrEmpty(idx))
                        AddPlaceholderTextStyle("idx:" + idx, listStyle, txBody);

                    if (!string.IsNullOrEmpty(type))
                        AddPlaceholderTextStyle("type:" + type, listStyle, txBody);

                    if (string.IsNullOrEmpty(idx) && string.IsNullOrEmpty(type))
                        AddPlaceholderTextStyle("type:body", listStyle, txBody);
                }
            }
        }

        private static void ReadPresentationDefaultTextStyle(
            XmlDocument document)
        {
            if (document == null ||
                activePresentationDefaultTextStyle == null)
            {
                return;
            }

            XmlNode defaultTextStyle =
                FindFirst(
                    document,
                    "defaultTextStyle");

            AddMasterStyleNode(
                activePresentationDefaultTextStyle,
                defaultTextStyle);
        }

        private static void ReadMasterTextStyles(XmlDocument document)
        {
            XmlNode txStyles = FindFirst(document, "txStyles");
            if (txStyles == null)
                return;

            AddMasterStyleNode(activeMasterTitleStyle, DirectChild(txStyles, "titleStyle"));
            AddMasterStyleNode(activeMasterBodyStyle, DirectChild(txStyles, "bodyStyle"));
            AddMasterStyleNode(activeMasterOtherStyle, DirectChild(txStyles, "otherStyle"));
        }

        private static void AddMasterStyleNode(
            PlaceholderTextStyleContext context,
            XmlNode styleNode)
        {
            if (context == null || styleNode == null)
                return;

            for (int level = 0; level < 9; level++)
            {
                XmlNode pPr = DirectChild(styleNode, "lvl" + (level + 1).ToString() + "pPr");
                if (pPr != null)
                    context.Levels[level].Add(pPr.CloneNode(true));
            }
        }

        private static void AddPlaceholderTextStyle(
            string key,
            XmlNode listStyle,
            XmlNode txBody)
        {
            if (string.IsNullOrEmpty(key) || activePlaceholderTextStyles == null)
                return;

            PlaceholderTextStyleContext context;
            if (!activePlaceholderTextStyles.TryGetValue(key, out context))
            {
                context = new PlaceholderTextStyleContext();
                activePlaceholderTextStyles.Add(key, context);
            }

            if (txBody != null)
            {
                XmlNode bodyProperties =
                    DirectChild(
                        txBody,
                        "bodyPr");

                if (bodyProperties != null)
                {
                    context.BodyProperties.Add(
                        bodyProperties.CloneNode(true));
                }
            }

            if (listStyle != null)
            {
                for (int level = 0; level < 9; level++)
                {
                    XmlNode pPr = DirectChild(listStyle, "lvl" + (level + 1).ToString() + "pPr");
                    if (pPr != null)
                        context.Levels[level].Add(pPr.CloneNode(true));
                }
            }

            // Some layouts put paragraph defaults in the placeholder's own
            // first paragraph instead of lstStyle.
            if (txBody != null)
            {
                foreach (XmlNode child in txBody.ChildNodes)
                {
                    if (child.LocalName != "p")
                        continue;

                    XmlNode pPr = DirectChild(child, "pPr");
                    if (pPr != null)
                        context.Levels[0].Add(pPr.CloneNode(true));
                    break;
                }
            }
        }

        private static XmlNode BuildRichInheritedShape(
            XmlNode shape)
        {
            if (shape == null)
                return shape;

            XmlNode clone =
                shape.CloneNode(true);

            XmlNode txBody =
                DirectChild(
                    clone,
                    "txBody");

            if (txBody == null)
                return clone;

            MergeInheritedTextBodyProperties(
                clone,
                txBody);

            XmlNode listStyle =
                DirectChild(
                    txBody,
                    "lstStyle");

            foreach (XmlNode paragraph in
                txBody.ChildNodes)
            {
                if (paragraph.LocalName != "p")
                    continue;

                XmlNode localPPr =
                    DirectChild(
                        paragraph,
                        "pPr");

                int level =
                    (int)Math.Max(
                        0,
                        Math.Min(
                            8,
                            GetLong(
                                localPPr,
                                "lvl",
                                0)));

                List<XmlNode> inherited =
                    GetInheritedRichParagraphProperties(
                        clone,
                        paragraph);

                XmlNode listStyleLevel =
                    listStyle != null
                        ? DirectChild(
                            listStyle,
                            "lvl" +
                            (level + 1).ToString() +
                            "pPr")
                        : null;

                if (inherited.Count == 0 &&
                    listStyleLevel == null &&
                    localPPr == null)
                {
                    continue;
                }

                XmlElement merged =
                    paragraph.OwnerDocument.CreateElement(
                        "a",
                        "pPr",
                        "http://schemas.openxmlformats.org/drawingml/2006/main");

                for (int i = 0;
                     i < inherited.Count;
                     i++)
                {
                    MergeRichParagraphProperties(
                        merged,
                        inherited[i]);
                }

                if (listStyleLevel != null)
                {
                    MergeRichParagraphProperties(
                        merged,
                        listStyleLevel);
                }

                if (localPPr != null)
                {
                    MergeRichParagraphProperties(
                        merged,
                        localPPr);
                }

                if (localPPr != null)
                {
                    paragraph.ReplaceChild(
                        merged,
                        localPPr);
                }
                else if (paragraph.FirstChild != null)
                {
                    paragraph.InsertBefore(
                        merged,
                        paragraph.FirstChild);
                }
                else
                {
                    paragraph.AppendChild(
                        merged);
                }
            }

            return clone;
        }

        private static void MergeInheritedTextBodyProperties(
            XmlNode shape,
            XmlNode txBody)
        {
            if (shape == null ||
                txBody == null)
            {
                return;
            }

            List<XmlNode> inherited =
                GetInheritedTextBodyProperties(
                    shape);

            XmlNode localBody =
                DirectChild(
                    txBody,
                    "bodyPr");

            if (inherited.Count == 0 &&
                localBody == null)
            {
                return;
            }

            XmlElement merged =
                txBody.OwnerDocument.CreateElement(
                    "a",
                    "bodyPr",
                    "http://schemas.openxmlformats.org/drawingml/2006/main");

            for (int i = 0;
                 i < inherited.Count;
                 i++)
            {
                MergeRichBodyProperties(
                    merged,
                    inherited[i]);
            }

            if (localBody != null)
            {
                MergeRichBodyProperties(
                    merged,
                    localBody);

                txBody.ReplaceChild(
                    merged,
                    localBody);
            }
            else if (txBody.FirstChild != null)
            {
                txBody.InsertBefore(
                    merged,
                    txBody.FirstChild);
            }
            else
            {
                txBody.AppendChild(
                    merged);
            }
        }

        private static List<XmlNode> GetInheritedTextBodyProperties(
            XmlNode shape)
        {
            List<XmlNode> result =
                new List<XmlNode>();

            XmlNode placeholder =
                FindFirst(
                    shape,
                    "ph");

            if (placeholder == null ||
                activePlaceholderTextStyles == null)
            {
                return result;
            }

            string type =
                GetAttr(
                    placeholder,
                    "type");
            string idx =
                GetAttr(
                    placeholder,
                    "idx");

            PlaceholderTextStyleContext specific =
                null;

            if (!string.IsNullOrEmpty(idx))
            {
                activePlaceholderTextStyles.TryGetValue(
                    "idx:" + idx,
                    out specific);
            }

            if (specific == null &&
                !string.IsNullOrEmpty(type))
            {
                activePlaceholderTextStyles.TryGetValue(
                    "type:" + type,
                    out specific);
            }

            if (specific == null)
            {
                activePlaceholderTextStyles.TryGetValue(
                    "type:body",
                    out specific);
            }

            if (specific != null)
            {
                for (int i = 0;
                     i < specific.BodyProperties.Count;
                     i++)
                {
                    result.Add(
                        specific.BodyProperties[i]);
                }
            }

            return result;
        }

        private static void MergeRichBodyProperties(
            XmlElement target,
            XmlNode source)
        {
            if (target == null ||
                source == null)
            {
                return;
            }

            if (source.Attributes != null)
            {
                foreach (XmlAttribute attribute in
                    source.Attributes)
                {
                    if (attribute == null ||
                        attribute.Prefix == "xmlns" ||
                        attribute.Name == "xmlns")
                    {
                        continue;
                    }

                    XmlAttribute copy =
                        target.OwnerDocument.CreateAttribute(
                            attribute.Prefix,
                            attribute.LocalName,
                            attribute.NamespaceURI ??
                            string.Empty);

                    copy.Value =
                        attribute.Value;

                    target.Attributes.SetNamedItem(
                        copy);
                }
            }

            foreach (XmlNode child in
                source.ChildNodes)
            {
                if (child.NodeType !=
                    XmlNodeType.Element)
                {
                    continue;
                }

                XmlNode existing =
                    DirectChild(
                        target,
                        child.LocalName);

                XmlNode imported =
                    target.OwnerDocument.ImportNode(
                        child,
                        true);

                if (existing != null)
                {
                    target.ReplaceChild(
                        imported,
                        existing);
                }
                else
                {
                    target.AppendChild(
                        imported);
                }
            }
        }

        private static void MergeRichParagraphProperties(
            XmlElement target,
            XmlNode source)
        {
            if (target == null || source == null)
                return;

            if (source.Attributes != null)
            {
                foreach (XmlAttribute attribute in source.Attributes)
                {
                    if (attribute == null || attribute.Prefix == "xmlns" || attribute.Name == "xmlns")
                        continue;

                    string namespaceUri = attribute.NamespaceURI ?? "";
                    XmlAttribute copy = target.OwnerDocument.CreateAttribute(
                        attribute.Prefix,
                        attribute.LocalName,
                        namespaceUri);
                    copy.Value = attribute.Value;
                    target.Attributes.SetNamedItem(copy);
                }
            }

            foreach (XmlNode child in source.ChildNodes)
            {
                if (child.NodeType != XmlNodeType.Element)
                    continue;

                XmlNode existing = DirectChild(target, child.LocalName);

                if (existing != null && child.LocalName == "defRPr")
                {
                    MergeRichRunPropertyNode((XmlElement)existing, child);
                    continue;
                }

                XmlNode imported = target.OwnerDocument.ImportNode(child, true);

                if (existing != null)
                    target.ReplaceChild(imported, existing);
                else
                    target.AppendChild(imported);
            }
        }

        private static void MergeRichRunPropertyNode(
            XmlElement target,
            XmlNode source)
        {
            if (target == null || source == null)
                return;

            if (source.Attributes != null)
            {
                foreach (XmlAttribute attribute in source.Attributes)
                {
                    if (attribute == null || attribute.Prefix == "xmlns" || attribute.Name == "xmlns")
                        continue;

                    XmlAttribute copy = target.OwnerDocument.CreateAttribute(
                        attribute.Prefix,
                        attribute.LocalName,
                        attribute.NamespaceURI ?? "");
                    copy.Value = attribute.Value;
                    target.Attributes.SetNamedItem(copy);
                }
            }

            foreach (XmlNode child in source.ChildNodes)
            {
                if (child.NodeType != XmlNodeType.Element)
                    continue;

                XmlNode existing = DirectChild(target, child.LocalName);
                XmlNode imported = target.OwnerDocument.ImportNode(child, true);

                if (existing != null)
                    target.ReplaceChild(imported, existing);
                else
                    target.AppendChild(imported);
            }
        }

        private static List<XmlNode> GetInheritedRichParagraphProperties(
            XmlNode shape,
            XmlNode paragraph)
        {
            List<XmlNode> result =
                new List<XmlNode>();

            XmlNode localPPr =
                paragraph != null
                    ? DirectChild(
                        paragraph,
                        "pPr")
                    : null;

            int level =
                (int)Math.Max(
                    0,
                    Math.Min(
                        8,
                        GetLong(
                            localPPr,
                            "lvl",
                            0)));

            AppendTextStyleLevel(
                result,
                activePresentationDefaultTextStyle,
                level);

            XmlNode placeholder =
                FindFirst(
                    shape,
                    "ph");

            if (placeholder == null)
                return result;

            string type =
                GetAttr(
                    placeholder,
                    "type");
            string idx =
                GetAttr(
                    placeholder,
                    "idx");

            PlaceholderTextStyleContext master =
                SelectMasterTextStyle(
                    type);

            AppendTextStyleLevel(
                result,
                master,
                level);

            if (activePlaceholderTextStyles != null)
            {
                PlaceholderTextStyleContext specific =
                    null;

                if (!string.IsNullOrEmpty(idx))
                {
                    activePlaceholderTextStyles.TryGetValue(
                        "idx:" + idx,
                        out specific);
                }

                if (specific == null &&
                    !string.IsNullOrEmpty(type))
                {
                    activePlaceholderTextStyles.TryGetValue(
                        "type:" + type,
                        out specific);
                }

                if (specific == null)
                {
                    activePlaceholderTextStyles.TryGetValue(
                        "type:body",
                        out specific);
                }

                AppendTextStyleLevel(
                    result,
                    specific,
                    level);
            }

            return result;
        }

        private static PlaceholderTextStyleContext SelectMasterTextStyle(string type)
        {
            if (type == "title" || type == "ctrTitle")
                return activeMasterTitleStyle;

            if (type == "body" || type == "obj" || type == "subTitle")
                return activeMasterBodyStyle;

            return activeMasterOtherStyle;
        }

        private static void AppendTextStyleLevel(
            List<XmlNode> target,
            PlaceholderTextStyleContext context,
            int level)
        {
            if (target == null || context == null)
                return;

            level =
                Math.Max(
                    0,
                    Math.Min(
                        8,
                        level));

            List<XmlNode> source =
                context.Levels[level];

            if (source.Count == 0)
            {
                for (int fallback = level - 1;
                     fallback >= 0;
                     fallback--)
                {
                    if (context.Levels[fallback].Count > 0)
                    {
                        source =
                            context.Levels[fallback];
                        break;
                    }
                }
            }

            for (int i = 0;
                 i < source.Count;
                 i++)
            {
                target.Add(
                    source[i]);
            }
        }
    }
}
