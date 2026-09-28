using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal static class DocxWriter
    {
        private const string MainContentType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        private const string StylesContentType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        private const string CorePropertiesContentType =
            "application/vnd.openxmlformats-package.core-properties+xml";
        private const string ExtendedPropertiesContentType =
            "application/vnd.openxmlformats-officedocument.extended-properties+xml";

        private const string OfficeDocumentRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string CorePropertiesRelationship =
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string ExtendedPropertiesRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        private const string StylesRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

        public static void Save(TextDocument document, string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path is required.", "outputPath");

            string destination = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(destination);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string stage = destination + ".writing";
            string backup = destination + ".backup";
            DeleteIfExists(stage);
            DeleteIfExists(backup);

            try
            {
                using (ZipArchive archive = ZipFile.Open(stage, ZipArchiveMode.Create))
                {
                    WriteContentTypes(archive);
                    WritePackageRelationships(archive);
                    WriteCoreProperties(archive, document.Title);
                    WriteExtendedProperties(archive);
                    WriteDocumentRelationships(archive);
                    WriteStyles(archive);
                    WriteDocument(archive, document);
                }

                ReplaceSafely(stage, destination, backup);
            }
            catch
            {
                DeleteIfExists(stage);
                throw;
            }
        }

        private static void WriteContentTypes(ZipArchive archive)
        {
            Dictionary<string, string> defaults = new Dictionary<string, string>();
            defaults.Add("rels", "application/vnd.openxmlformats-package.relationships+xml");
            defaults.Add("xml", "application/xml");

            Dictionary<string, string> overrides = new Dictionary<string, string>();
            overrides.Add("word/document.xml", MainContentType);
            overrides.Add("word/styles.xml", StylesContentType);
            overrides.Add("docProps/core.xml", CorePropertiesContentType);
            overrides.Add("docProps/app.xml", ExtendedPropertiesContentType);

            XmlDocument document =
                OpcPackageUtility.CreateContentTypesDocument(defaults, overrides);
            OpcPackageUtility.WriteXmlPart(
                archive,
                "[Content_Types].xml",
                document);
        }

        private static void WritePackageRelationships(ZipArchive archive)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship(
                "rId1",
                OfficeDocumentRelationship,
                "word/document.xml"));
            relationships.Add(Relationship(
                "rId2",
                CorePropertiesRelationship,
                "docProps/core.xml"));
            relationships.Add(Relationship(
                "rId3",
                ExtendedPropertiesRelationship,
                "docProps/app.xml"));

            OpcPackageUtility.WriteXmlPart(
                archive,
                "_rels/.rels",
                OpcPackageUtility.CreateRelationshipsDocument(relationships));
        }

        private static void WriteDocumentRelationships(ZipArchive archive)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship(
                "rId1",
                StylesRelationship,
                "styles.xml"));

            OpcPackageUtility.WriteXmlPart(
                archive,
                "word/_rels/document.xml.rels",
                OpcPackageUtility.CreateRelationshipsDocument(relationships));
        }

        private static void WriteStyles(ZipArchive archive)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:docDefaults>" +
                "<w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault>" +
                "<w:pPrDefault><w:pPr/></w:pPrDefault>" +
                "</w:docDefaults>" +
                "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
                "<w:name w:val=\"Normal\"/><w:qFormat/>" +
                "</w:style>" +
                "</w:styles>";

            WriteXmlString(archive, "word/styles.xml", xml);
        }

        private static void WriteDocument(
            ZipArchive archive,
            TextDocument document)
        {
            StringBuilder body = new StringBuilder();

            for (int i = 0; i < document.Paragraphs.Count; i++)
            {
                DocumentParagraph paragraph = document.Paragraphs[i];
                if (paragraph == null)
                    continue;

                body.Append(BuildParagraph(paragraph));
            }

            body.Append(
                "<w:sectPr>" +
                "<w:pgSz w:w=\"11906\" w:h=\"16838\"/>" +
                "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" " +
                "w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
                "</w:sectPr>");

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<w:body>" + body.ToString() + "</w:body>" +
                "</w:document>";

            WriteXmlString(archive, "word/document.xml", xml);
        }

        private static string BuildParagraph(DocumentParagraph paragraph)
        {
            StringBuilder xml = new StringBuilder();
            xml.Append("<w:p>");
            xml.Append("<w:pPr><w:jc w:val=\"");
            xml.Append(AlignmentValue(paragraph.Alignment));
            xml.Append("\"/></w:pPr>");

            for (int i = 0; i < paragraph.Runs.Count; i++)
            {
                DocumentTextRun run = paragraph.Runs[i];
                if (run == null)
                    continue;

                xml.Append(BuildRun(run));
            }

            xml.Append("</w:p>");
            return xml.ToString();
        }

        private static string BuildRun(DocumentTextRun run)
        {
            string family = EscapeXml(
                string.IsNullOrEmpty(run.FontFamily)
                    ? "Arial"
                    : run.FontFamily);
            int halfPoints = (int)Math.Round(
                Math.Max(1f, Math.Min(400f, run.FontSizePoints)) * 2.0f);

            StringBuilder properties = new StringBuilder();
            properties.Append("<w:rPr>");
            properties.Append("<w:rFonts w:ascii=\"");
            properties.Append(family);
            properties.Append("\" w:hAnsi=\"");
            properties.Append(family);
            properties.Append("\" w:eastAsia=\"");
            properties.Append(family);
            properties.Append("\"/>");

            if (run.Bold)
                properties.Append("<w:b/>");
            if (run.Italic)
                properties.Append("<w:i/>");
            if (run.Underline)
                properties.Append("<w:u w:val=\"single\"/>");

            properties.Append("<w:color w:val=\"");
            properties.Append(NormalizeColor(run.ColorHex));
            properties.Append("\"/>");
            properties.Append("<w:sz w:val=\"");
            properties.Append(halfPoints.ToString());
            properties.Append("\"/>");
            properties.Append("<w:szCs w:val=\"");
            properties.Append(halfPoints.ToString());
            properties.Append("\"/>");
            properties.Append("</w:rPr>");

            string text = run.Text ?? string.Empty;
            string preserve =
                text.StartsWith(" ", StringComparison.Ordinal) ||
                text.EndsWith(" ", StringComparison.Ordinal)
                    ? " xml:space=\"preserve\""
                    : string.Empty;

            return
                "<w:r>" +
                properties.ToString() +
                "<w:t" + preserve + ">" +
                EscapeXml(text) +
                "</w:t></w:r>";
        }

        private static void WriteCoreProperties(
            ZipArchive archive,
            string title)
        {
            string safeTitle = EscapeXml(
                string.IsNullOrEmpty(title)
                    ? "New Document"
                    : title);
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                "<dc:title>" + safeTitle + "</dc:title>" +
                "<dc:creator>PowerPointLite</dc:creator>" +
                "<cp:lastModifiedBy>PowerPointLite</cp:lastModifiedBy>" +
                "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + now + "</dcterms:created>" +
                "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + now + "</dcterms:modified>" +
                "</cp:coreProperties>";

            WriteXmlString(archive, "docProps/core.xml", xml);
        }

        private static void WriteExtendedProperties(ZipArchive archive)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
                "<Application>PowerPointLite</Application>" +
                "<DocSecurity>0</DocSecurity>" +
                "<ScaleCrop>false</ScaleCrop>" +
                "<Company></Company>" +
                "<LinksUpToDate>false</LinksUpToDate>" +
                "<SharedDoc>false</SharedDoc>" +
                "<HyperlinksChanged>false</HyperlinksChanged>" +
                "<AppVersion>1.0</AppVersion>" +
                "</Properties>";

            WriteXmlString(archive, "docProps/app.xml", xml);
        }

        private static OpcRelationship Relationship(
            string id,
            string type,
            string target)
        {
            OpcRelationship relationship = new OpcRelationship();
            relationship.Id = id;
            relationship.Type = type;
            relationship.Target = target;
            return relationship;
        }

        private static void WriteXmlString(
            ZipArchive archive,
            string partName,
            string xml)
        {
            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = true;
            document.LoadXml(xml);
            OpcPackageUtility.WriteXmlPart(archive, partName, document);
        }

        private static string AlignmentValue(
            DocumentParagraphAlignment alignment)
        {
            if (alignment == DocumentParagraphAlignment.Center)
                return "center";
            if (alignment == DocumentParagraphAlignment.Right)
                return "right";
            if (alignment == DocumentParagraphAlignment.Justify)
                return "both";
            return "left";
        }

        private static string NormalizeColor(string value)
        {
            string candidate = (value ?? string.Empty)
                .Trim()
                .TrimStart('#')
                .ToUpperInvariant();

            if (candidate.Length != 6)
                return "20242A";

            for (int i = 0; i < candidate.Length; i++)
            {
                char ch = candidate[i];
                bool valid =
                    (ch >= '0' && ch <= '9') ||
                    (ch >= 'A' && ch <= 'F');
                if (!valid)
                    return "20242A";
            }

            return candidate;
        }

        private static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }

        private static void ReplaceSafely(
            string stage,
            string destination,
            string backup)
        {
            if (!File.Exists(destination))
            {
                File.Move(stage, destination);
                return;
            }

            File.Move(destination, backup);

            try
            {
                File.Move(stage, destination);
                DeleteIfExists(backup);
            }
            catch
            {
                try
                {
                    DeleteIfExists(destination);
                    if (File.Exists(backup))
                        File.Move(backup, destination);
                }
                catch { }

                throw;
            }
        }

        private static void DeleteIfExists(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch { }
        }
    }
}
