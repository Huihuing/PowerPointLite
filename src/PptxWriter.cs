using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Xml;

namespace PptxViewer
{
    internal static class PptxWriter
    {
        private const string PresentationContentType =
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
        private const string SlideContentType =
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        private const string SlideMasterContentType =
            "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        private const string SlideLayoutContentType =
            "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
        private const string ThemeContentType =
            "application/vnd.openxmlformats-officedocument.theme+xml";
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
        private const string SlideRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        private const string SlideMasterRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
        private const string SlideLayoutRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
        private const string ThemeRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

        public static void CreateNewPresentation(string outputPath, string title)
        {
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path is required.", "outputPath");

            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            using (ZipArchive archive = ZipFile.Open(outputPath, ZipArchiveMode.Create))
            {
                WriteContentTypes(archive);
                WritePackageRelationships(archive);
                WriteCoreProperties(archive, title);
                WriteExtendedProperties(archive);
                WritePresentation(archive);
                WritePresentationRelationships(archive);
                WriteSlideMaster(archive);
                WriteSlideMasterRelationships(archive);
                WriteSlideLayout(archive);
                WriteSlideLayoutRelationships(archive);
                WriteTheme(archive);
                WriteFirstSlide(archive, title);
                WriteSlideRelationships(archive);
            }
        }

        private static void WriteContentTypes(ZipArchive archive)
        {
            Dictionary<string, string> defaults = new Dictionary<string, string>();
            defaults.Add("rels", "application/vnd.openxmlformats-package.relationships+xml");
            defaults.Add("xml", "application/xml");

            Dictionary<string, string> overrides = new Dictionary<string, string>();
            overrides.Add("ppt/presentation.xml", PresentationContentType);
            overrides.Add("ppt/slides/slide1.xml", SlideContentType);
            overrides.Add("ppt/slideMasters/slideMaster1.xml", SlideMasterContentType);
            overrides.Add("ppt/slideLayouts/slideLayout1.xml", SlideLayoutContentType);
            overrides.Add("ppt/theme/theme1.xml", ThemeContentType);
            overrides.Add("docProps/core.xml", CorePropertiesContentType);
            overrides.Add("docProps/app.xml", ExtendedPropertiesContentType);

            XmlDocument document = OpcPackageUtility.CreateContentTypesDocument(defaults, overrides);
            OpcPackageUtility.WriteXmlPart(archive, "[Content_Types].xml", document);
        }

        private static void WritePackageRelationships(ZipArchive archive)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship("rId1", OfficeDocumentRelationship, "ppt/presentation.xml"));
            relationships.Add(Relationship("rId2", CorePropertiesRelationship, "docProps/core.xml"));
            relationships.Add(Relationship("rId3", ExtendedPropertiesRelationship, "docProps/app.xml"));

            XmlDocument document = OpcPackageUtility.CreateRelationshipsDocument(relationships);
            OpcPackageUtility.WriteXmlPart(archive, "_rels/.rels", document);
        }

        private static void WritePresentation(ZipArchive archive)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>" +
                "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId2\"/></p:sldIdLst>" +
                "<p:sldSz cx=\"12192000\" cy=\"6858000\" type=\"screen16x9\"/>" +
                "<p:notesSz cx=\"6858000\" cy=\"9144000\"/>" +
                "<p:defaultTextStyle/>" +
                "</p:presentation>";

            WriteXmlString(archive, "ppt/presentation.xml", xml);
        }

        private static void WritePresentationRelationships(ZipArchive archive)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship("rId1", SlideMasterRelationship, "slideMasters/slideMaster1.xml"));
            relationships.Add(Relationship("rId2", SlideRelationship, "slides/slide1.xml"));

            XmlDocument document = OpcPackageUtility.CreateRelationshipsDocument(relationships);
            OpcPackageUtility.WriteXmlPart(archive, "ppt/_rels/presentation.xml.rels", document);
        }

        private static void WriteSlideMaster(ZipArchive archive)
        {
            string spTree = EmptyShapeTreeXml();

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sldMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:cSld><p:spTree>" + spTree + "</p:spTree></p:cSld>" +
                "<p:clrMap accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" " +
                "accent5=\"accent5\" accent6=\"accent6\" bg1=\"lt1\" bg2=\"lt2\" hlink=\"hlink\" " +
                "folHlink=\"folHlink\" tx1=\"dk1\" tx2=\"dk2\"/>" +
                "<p:sldLayoutIdLst><p:sldLayoutId id=\"1\" r:id=\"rId2\"/></p:sldLayoutIdLst>" +
                "<p:txStyles>" +
                "<p:titleStyle><a:lvl1pPr algn=\"l\"><a:defRPr sz=\"3200\"/></a:lvl1pPr></p:titleStyle>" +
                "<p:bodyStyle><a:lvl1pPr marL=\"342900\" indent=\"-342900\"><a:defRPr sz=\"1800\"/></a:lvl1pPr></p:bodyStyle>" +
                "<p:otherStyle><a:defPPr><a:defRPr lang=\"en-US\"/></a:defPPr><a:lvl1pPr><a:defRPr sz=\"1800\"/></a:lvl1pPr></p:otherStyle>" +
                "</p:txStyles>" +
                "</p:sldMaster>";

            WriteXmlString(archive, "ppt/slideMasters/slideMaster1.xml", xml);
        }

        private static void WriteSlideMasterRelationships(ZipArchive archive)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship("rId1", ThemeRelationship, "../theme/theme1.xml"));
            relationships.Add(Relationship("rId2", SlideLayoutRelationship, "../slideLayouts/slideLayout1.xml"));

            XmlDocument document = OpcPackageUtility.CreateRelationshipsDocument(relationships);
            OpcPackageUtility.WriteXmlPart(archive, "ppt/slideMasters/_rels/slideMaster1.xml.rels", document);
        }

        private static void WriteSlideLayout(ZipArchive archive)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" type=\"blank\" preserve=\"1\">" +
                "<p:cSld name=\"Blank\"><p:spTree>" + EmptyShapeTreeXml() + "</p:spTree></p:cSld>" +
                "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" +
                "</p:sldLayout>";

            WriteXmlString(archive, "ppt/slideLayouts/slideLayout1.xml", xml);
        }

        private static void WriteSlideLayoutRelationships(ZipArchive archive)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship("rId1", SlideMasterRelationship, "../slideMasters/slideMaster1.xml"));

            XmlDocument document = OpcPackageUtility.CreateRelationshipsDocument(relationships);
            OpcPackageUtility.WriteXmlPart(archive, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", document);
        }

        private static void WriteFirstSlide(ZipArchive archive, string title)
        {
            string safeTitle = EscapeXml(string.IsNullOrEmpty(title) ? "New Presentation" : title);

            string textShape =
                "<p:sp>" +
                "<p:nvSpPr><p:cNvPr id=\"2\" name=\"Title\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>" +
                "<p:spPr>" +
                "<a:xfrm><a:off x=\"914400\" y=\"2057400\"/><a:ext cx=\"10363200\" cy=\"1371600\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>" +
                "</p:spPr>" +
                "<p:txBody><a:bodyPr anchor=\"ctr\"/><a:lstStyle/>" +
                "<a:p><a:pPr algn=\"ctr\"/>" +
                "<a:r><a:rPr lang=\"ko-KR\" sz=\"2800\" b=\"1\"><a:solidFill><a:srgbClr val=\"20242A\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr>" +
                "<a:t>" + safeTitle + "</a:t></a:r><a:endParaRPr lang=\"ko-KR\" sz=\"2800\"/></a:p>" +
                "</p:txBody>" +
                "</p:sp>";

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:cSld><p:spTree>" + EmptyShapeTreeXml() + textShape + "</p:spTree></p:cSld>" +
                "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" +
                "</p:sld>";

            WriteXmlString(archive, "ppt/slides/slide1.xml", xml);
        }

        private static void WriteSlideRelationships(ZipArchive archive)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship("rId1", SlideLayoutRelationship, "../slideLayouts/slideLayout1.xml"));

            XmlDocument document = OpcPackageUtility.CreateRelationshipsDocument(relationships);
            OpcPackageUtility.WriteXmlPart(archive, "ppt/slides/_rels/slide1.xml.rels", document);
        }

        private static void WriteTheme(ZipArchive archive)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"PowerPointLite Neutral\">" +
                "<a:themeElements>" +
                "<a:clrScheme name=\"PowerPointLite Neutral\">" +
                Color("dk1", "20242A") + Color("lt1", "FFFFFF") + Color("dk2", "44546A") + Color("lt2", "F2F3F5") +
                Color("accent1", "5B8CFF") + Color("accent2", "31B6A1") + Color("accent3", "F4A261") +
                Color("accent4", "9B7EDE") + Color("accent5", "E76F8A") + Color("accent6", "6C9A8B") +
                Color("hlink", "356AE6") + Color("folHlink", "7A5AA6") +
                "</a:clrScheme>" +
                "<a:fontScheme name=\"PowerPointLite Neutral\">" +
                "<a:majorFont><a:latin typeface=\"Arial\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
                "<a:minorFont><a:latin typeface=\"Arial\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>" +
                "</a:fontScheme>" +
                "<a:fmtScheme name=\"PowerPointLite Neutral\">" +
                "<a:fillStyleLst>" + SolidPhClr() + SolidPhClr() + SolidPhClr() + "</a:fillStyleLst>" +
                "<a:lnStyleLst>" + LineStyle("6350") + LineStyle("12700") + LineStyle("19050") + "</a:lnStyleLst>" +
                "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
                "<a:bgFillStyleLst>" + SolidPhClr() + SolidPhClr() + SolidPhClr() + "</a:bgFillStyleLst>" +
                "</a:fmtScheme>" +
                "</a:themeElements>" +
                "</a:theme>";

            WriteXmlString(archive, "ppt/theme/theme1.xml", xml);
        }

        private static void WriteCoreProperties(ZipArchive archive, string title)
        {
            string safeTitle = EscapeXml(string.IsNullOrEmpty(title) ? "New Presentation" : title);
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                "<dc:title>" + safeTitle + "</dc:title><dc:creator>PowerPointLite</dc:creator>" +
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
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>PowerPointLite</Application><PresentationFormat>Widescreen</PresentationFormat>" +
                "<Slides>1</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips>" +
                "<ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate>" +
                "<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion>" +
                "</Properties>";

            WriteXmlString(archive, "docProps/app.xml", xml);
        }

        private static OpcRelationship Relationship(string id, string type, string target)
        {
            OpcRelationship relationship = new OpcRelationship();
            relationship.Id = id;
            relationship.Type = type;
            relationship.Target = target;
            return relationship;
        }

        private static void WriteXmlString(ZipArchive archive, string partName, string xml)
        {
            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = true;
            document.LoadXml(xml);
            OpcPackageUtility.WriteXmlPart(archive, partName, document);
        }

        private static string EmptyShapeTreeXml()
        {
            return
                "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
                "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>" +
                "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>";
        }

        private static string Color(string name, string value)
        {
            return "<a:" + name + "><a:srgbClr val=\"" + value + "\"/></a:" + name + ">";
        }

        private static string SolidPhClr()
        {
            return "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
        }

        private static string LineStyle(string width)
        {
            return "<a:ln w=\"" + width + "\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" +
                   "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>";
        }

        private static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }
    }
}
