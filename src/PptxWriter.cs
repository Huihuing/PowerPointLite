using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
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
        private const string ImageRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        public static void CreateNewPresentation(string outputPath, string title)
        {
            PresentationDocument document = PresentationDocument.CreateNew(title);
            Save(document, outputPath);
        }

        public static void Save(PresentationDocument document, string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path is required.", "outputPath");

            if (document.Slides == null || document.Slides.Count == 0)
                throw new InvalidOperationException("A presentation must contain at least one slide.");

            string fullPath = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string temporaryPath = fullPath + ".writing";

            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);

            try
            {
                using (ZipArchive archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Create))
                {
                    WriteContentTypes(archive, document.Slides.Count);
                    WritePackageRelationships(archive);
                    WriteCoreProperties(archive, document.Title);
                    WriteExtendedProperties(archive, document.Slides.Count);
                    WritePresentation(archive, document);
                    WritePresentationRelationships(archive, document.Slides.Count);
                    WriteSlideMaster(archive);
                    WriteSlideMasterRelationships(archive);
                    WriteSlideLayout(archive);
                    WriteSlideLayoutRelationships(archive);
                    WriteTheme(archive);

                    for (int i = 0; i < document.Slides.Count; i++)
                    {
                        int slideNumber = i + 1;
                        PresentationSlide slide = document.Slides[i];
                        WriteSlide(archive, slideNumber, slide);
                        WriteSlideMedia(archive, slideNumber, slide);
                        WriteSlideRelationships(archive, slideNumber, slide);
                    }
                }

                ReplaceDestinationSafely(temporaryPath, fullPath);
            }
            catch
            {
                try
                {
                    if (File.Exists(temporaryPath))
                        File.Delete(temporaryPath);
                }
                catch { }

                throw;
            }
        }

        private static void ReplaceDestinationSafely(
            string temporaryPath,
            string destinationPath)
        {
            if (!File.Exists(destinationPath))
            {
                File.Move(temporaryPath, destinationPath);
                return;
            }

            string backupPath = destinationPath + ".backup";

            if (File.Exists(backupPath))
                File.Delete(backupPath);

            File.Move(destinationPath, backupPath);

            try
            {
                File.Move(temporaryPath, destinationPath);
                File.Delete(backupPath);
            }
            catch
            {
                try
                {
                    if (File.Exists(destinationPath))
                        File.Delete(destinationPath);

                    if (File.Exists(backupPath))
                        File.Move(backupPath, destinationPath);
                }
                catch { }

                throw;
            }
        }

        private static void WriteContentTypes(ZipArchive archive, int slideCount)
        {
            Dictionary<string, string> defaults = new Dictionary<string, string>();
            defaults.Add("rels", "application/vnd.openxmlformats-package.relationships+xml");
            defaults.Add("xml", "application/xml");
            defaults.Add("png", "image/png");
            defaults.Add("jpg", "image/jpeg");
            defaults.Add("jpeg", "image/jpeg");
            defaults.Add("gif", "image/gif");
            defaults.Add("bmp", "image/bmp");

            Dictionary<string, string> overrides = new Dictionary<string, string>();
            overrides.Add("ppt/presentation.xml", PresentationContentType);
            overrides.Add("ppt/slideMasters/slideMaster1.xml", SlideMasterContentType);
            overrides.Add("ppt/slideLayouts/slideLayout1.xml", SlideLayoutContentType);
            overrides.Add("ppt/theme/theme1.xml", ThemeContentType);
            overrides.Add("docProps/core.xml", CorePropertiesContentType);
            overrides.Add("docProps/app.xml", ExtendedPropertiesContentType);

            for (int i = 1; i <= slideCount; i++)
                overrides.Add("ppt/slides/slide" + i.ToString() + ".xml", SlideContentType);

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

        private static void WritePresentation(ZipArchive archive, PresentationDocument document)
        {
            StringBuilder slideIds = new StringBuilder();

            for (int i = 0; i < document.Slides.Count; i++)
            {
                slideIds.Append("<p:sldId id=\"");
                slideIds.Append((256 + i).ToString());
                slideIds.Append("\" r:id=\"rId");
                slideIds.Append((i + 2).ToString());
                slideIds.Append("\"/>");
            }

            long width = document.WidthEmu > 0
                ? document.WidthEmu
                : PresentationDocument.DefaultWidthEmu;
            long height = document.HeightEmu > 0
                ? document.HeightEmu
                : PresentationDocument.DefaultHeightEmu;

            string sizeType =
                width == PresentationDocument.DefaultWidthEmu &&
                height == PresentationDocument.DefaultHeightEmu
                    ? " type=\"screen16x9\""
                    : string.Empty;

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>" +
                "<p:sldIdLst>" + slideIds.ToString() + "</p:sldIdLst>" +
                "<p:sldSz cx=\"" + width.ToString() + "\" cy=\"" + height.ToString() + "\"" + sizeType + "/>" +
                "<p:notesSz cx=\"6858000\" cy=\"9144000\"/>" +
                "<p:defaultTextStyle/>" +
                "</p:presentation>";

            WriteXmlString(archive, "ppt/presentation.xml", xml);
        }

        private static void WritePresentationRelationships(ZipArchive archive, int slideCount)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship("rId1", SlideMasterRelationship, "slideMasters/slideMaster1.xml"));

            for (int i = 1; i <= slideCount; i++)
            {
                relationships.Add(Relationship(
                    "rId" + (i + 1).ToString(),
                    SlideRelationship,
                    "slides/slide" + i.ToString() + ".xml"));
            }

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

        private static void WriteSlide(
            ZipArchive archive,
            int slideNumber,
            PresentationSlide slide)
        {
            if (slide == null)
                slide = new PresentationSlide();

            StringBuilder objects = new StringBuilder();
            int shapeId = 2;

            slide.SynchronizeObjectOrder();

            for (int i = 0;
                 i < slide.ObjectOrder.Count;
                 i++)
            {
                PresentationLayerEntry entry =
                    slide.ObjectOrder[i];

                if (entry == null)
                    continue;

                if (entry.Kind ==
                        PresentationLayerKind.Shape &&
                    entry.Index >= 0 &&
                    entry.Index < slide.Shapes.Count)
                {
                    objects.Append(
                        BuildBasicShape(
                            slide.Shapes[entry.Index],
                            shapeId++));
                }
                else if (entry.Kind ==
                             PresentationLayerKind.Image &&
                         entry.Index >= 0 &&
                         entry.Index < slide.Images.Count)
                {
                    objects.Append(
                        BuildImageShape(
                            slide.Images[entry.Index],
                            shapeId++,
                            "rId" +
                            (entry.Index + 2).ToString()));
                }
                else if (entry.Kind ==
                             PresentationLayerKind.TextBox &&
                         entry.Index >= 0 &&
                         entry.Index < slide.TextBoxes.Count)
                {
                    objects.Append(
                        BuildTextShape(
                            slide.TextBoxes[entry.Index],
                            shapeId++));
                }
            }

            string slideName = EscapeXml(
                string.IsNullOrEmpty(slide.Name)
                    ? "Slide " + slideNumber.ToString()
                    : slide.Name);

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                "<p:cSld name=\"" + slideName + "\"><p:spTree>" +
                EmptyShapeTreeXml() + objects.ToString() +
                "</p:spTree></p:cSld>" +
                "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" +
                "</p:sld>";

            WriteXmlString(
                archive,
                "ppt/slides/slide" + slideNumber.ToString() + ".xml",
                xml);
        }

        private static string BuildBasicShape(PresentationShape shape, int shapeId)
        {
            if (shape == null)
                shape = new PresentationShape();

            long x = Math.Max(0L, shape.X);
            long y = Math.Max(0L, shape.Y);
            long width = Math.Max(1L, shape.Width);
            long height = Math.Max(1L, shape.Height);
            long lineWidth = (long)Math.Round(
                Math.Max(0.25f, Math.Min(20f, shape.LineWidthPoints)) * 12700.0f);

            string name = EscapeXml(
                string.IsNullOrEmpty(shape.Name)
                    ? "Shape " + shapeId.ToString()
                    : shape.Name);

            return
                "<p:sp>" +
                "<p:nvSpPr><p:cNvPr id=\"" + shapeId.ToString() + "\" name=\"" + name + "\"/>" +
                "<p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
                "<p:spPr>" +
                "<a:xfrm><a:off x=\"" + x.ToString() + "\" y=\"" + y.ToString() + "\"/>" +
                "<a:ext cx=\"" + width.ToString() + "\" cy=\"" + height.ToString() + "\"/></a:xfrm>" +
                "<a:prstGeom prst=\"" + ShapePreset(shape.Kind) + "\"><a:avLst/></a:prstGeom>" +
                "<a:solidFill><a:srgbClr val=\"" + NormalizeColor(shape.FillColorHex) + "\"/></a:solidFill>" +
                "<a:ln w=\"" + lineWidth.ToString() + "\"><a:solidFill><a:srgbClr val=\"" +
                NormalizeColor(shape.LineColorHex) + "\"/></a:solidFill></a:ln>" +
                "</p:spPr>" +
                "</p:sp>";
        }

        private static string BuildImageShape(
            PresentationImage image,
            int shapeId,
            string relationshipId)
        {
            if (image == null)
                image = new PresentationImage();

            long x = Math.Max(0L, image.X);
            long y = Math.Max(0L, image.Y);
            long width = Math.Max(1L, image.Width);
            long height = Math.Max(1L, image.Height);
            string name = EscapeXml(
                string.IsNullOrEmpty(image.Name)
                    ? "Image " + shapeId.ToString()
                    : image.Name);

            int cropLeft;
            int cropTop;
            int cropRight;
            int cropBottom;
            PresentationRenderPrimitives.NormalizeCrop(
                image.CropLeft,
                image.CropTop,
                image.CropRight,
                image.CropBottom,
                out cropLeft,
                out cropTop,
                out cropRight,
                out cropBottom);

            string cropXml =
                cropLeft == 0 &&
                cropTop == 0 &&
                cropRight == 0 &&
                cropBottom == 0
                    ? string.Empty
                    : "<a:srcRect l=\"" +
                      cropLeft.ToString() +
                      "\" t=\"" +
                      cropTop.ToString() +
                      "\" r=\"" +
                      cropRight.ToString() +
                      "\" b=\"" +
                      cropBottom.ToString() +
                      "\"/>";

            int rotationUnits =
                PresentationRenderPrimitives
                    .NormalizeRotationUnits(
                        image.RotationUnits);
            int opacity =
                PresentationRenderPrimitives
                    .ClampOpacity(
                        image.Opacity);
            string transformAttributes =
                (rotationUnits == 0
                    ? string.Empty
                    : " rot=\"" +
                      rotationUnits.ToString() +
                      "\"") +
                (image.FlipHorizontal
                    ? " flipH=\"1\""
                    : string.Empty) +
                (image.FlipVertical
                    ? " flipV=\"1\""
                    : string.Empty);
            string alphaXml =
                opacity >= 100000
                    ? string.Empty
                    : "<a:alphaModFix amt=\"" +
                      opacity.ToString() +
                      "\"/>";

            return
                "<p:pic>" +
                "<p:nvPicPr><p:cNvPr id=\"" + shapeId.ToString() + "\" name=\"" + name + "\"/>" +
                "<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>" +
                "<p:blipFill><a:blip r:embed=\"" + EscapeXml(relationshipId) + "\">" +
                alphaXml +
                "</a:blip>" +
                cropXml +
                "<a:stretch><a:fillRect/></a:stretch></p:blipFill>" +
                "<p:spPr><a:xfrm" + transformAttributes + "><a:off x=\"" + x.ToString() + "\" y=\"" + y.ToString() + "\"/>" +
                "<a:ext cx=\"" + width.ToString() + "\" cy=\"" + height.ToString() + "\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>" +
                "</p:pic>";
        }

        private static string BuildTextShape(PresentationTextBox box, int shapeId)
        {
            if (box == null)
                box = new PresentationTextBox();

            long x = Math.Max(0L, box.X);
            long y = Math.Max(0L, box.Y);
            long width = Math.Max(1L, box.Width);
            long height = Math.Max(1L, box.Height);

            int fontSize = (int)Math.Round(box.FontSizePoints * 100.0f);
            fontSize = Math.Max(100, Math.Min(40000, fontSize));

            string fontFamily = EscapeXml(
                string.IsNullOrEmpty(box.FontFamily)
                    ? "Arial"
                    : box.FontFamily);

            string color = NormalizeColor(box.ColorHex);
            string alignment = AlignmentValue(box.Alignment);
            string bold = box.Bold ? " b=\"1\"" : string.Empty;
            string italic = box.Italic ? " i=\"1\"" : string.Empty;
            string name = EscapeXml(
                string.IsNullOrEmpty(box.Name)
                    ? "Text Box " + shapeId.ToString()
                    : box.Name);

            string runProperties =
                "<a:rPr lang=\"ko-KR\" sz=\"" + fontSize.ToString() + "\"" +
                bold + italic + ">" +
                "<a:solidFill><a:srgbClr val=\"" + color + "\"/></a:solidFill>" +
                "<a:latin typeface=\"" + fontFamily + "\"/>" +
                "<a:ea typeface=\"" + fontFamily + "\"/>" +
                "</a:rPr>";

            string paragraphXml = BuildParagraphs(
                box.Text,
                alignment,
                fontSize,
                runProperties);

            return
                "<p:sp>" +
                "<p:nvSpPr><p:cNvPr id=\"" + shapeId.ToString() + "\" name=\"" + name + "\"/>" +
                "<p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>" +
                "<p:spPr>" +
                "<a:xfrm><a:off x=\"" + x.ToString() + "\" y=\"" + y.ToString() + "\"/>" +
                "<a:ext cx=\"" + width.ToString() + "\" cy=\"" + height.ToString() + "\"/></a:xfrm>" +
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" +
                "<a:noFill/><a:ln><a:noFill/></a:ln>" +
                "</p:spPr>" +
                "<p:txBody><a:bodyPr wrap=\"square\" anchor=\"ctr\"/><a:lstStyle/>" +
                paragraphXml +
                "</p:txBody>" +
                "</p:sp>";
        }

        private static string BuildParagraphs(
            string text,
            string alignment,
            int fontSize,
            string runProperties)
        {
            string normalized = (text ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');

            string[] lines = normalized.Split('\n');
            StringBuilder result = new StringBuilder();

            if (lines.Length == 0)
                lines = new string[] { string.Empty };

            for (int i = 0; i < lines.Length; i++)
            {
                result.Append("<a:p><a:pPr algn=\"");
                result.Append(alignment);
                result.Append("\"/>");

                if (!string.IsNullOrEmpty(lines[i]))
                {
                    result.Append("<a:r>");
                    result.Append(runProperties);
                    result.Append("<a:t>");
                    result.Append(EscapeXml(lines[i]));
                    result.Append("</a:t></a:r>");
                }

                result.Append("<a:endParaRPr lang=\"ko-KR\" sz=\"");
                result.Append(fontSize.ToString());
                result.Append("\"/></a:p>");
            }

            return result.ToString();
        }

        private static void WriteSlideMedia(
            ZipArchive archive,
            int slideNumber,
            PresentationSlide slide)
        {
            if (slide == null)
                return;

            for (int i = 0; i < slide.Images.Count; i++)
            {
                PresentationImage image = slide.Images[i];
                if (image == null || image.Data == null || image.Data.Length == 0)
                    continue;

                string extension = NormalizeImageExtension(image.Extension);
                string partName = GetImagePartName(slideNumber, i + 1, extension);
                ZipArchiveEntry entry = archive.CreateEntry(partName, CompressionLevel.Optimal);

                using (Stream stream = entry.Open())
                    stream.Write(image.Data, 0, image.Data.Length);
            }
        }

        private static void WriteSlideRelationships(
            ZipArchive archive,
            int slideNumber,
            PresentationSlide slide)
        {
            List<OpcRelationship> relationships = new List<OpcRelationship>();
            relationships.Add(Relationship("rId1", SlideLayoutRelationship, "../slideLayouts/slideLayout1.xml"));

            if (slide != null)
            {
                for (int i = 0; i < slide.Images.Count; i++)
                {
                    PresentationImage image = slide.Images[i];
                    string extension = NormalizeImageExtension(
                        image == null ? string.Empty : image.Extension);

                    relationships.Add(Relationship(
                        "rId" + (i + 2).ToString(),
                        ImageRelationship,
                        "../media/" + Path.GetFileName(
                            GetImagePartName(slideNumber, i + 1, extension))));
                }
            }

            XmlDocument document = OpcPackageUtility.CreateRelationshipsDocument(relationships);
            OpcPackageUtility.WriteXmlPart(
                archive,
                "ppt/slides/_rels/slide" + slideNumber.ToString() + ".xml.rels",
                document);
        }

        private static string GetImagePartName(
            int slideNumber,
            int imageNumber,
            string extension)
        {
            return "ppt/media/slide" + slideNumber.ToString() +
                "_image" + imageNumber.ToString() + "." + extension;
        }

        private static string NormalizeImageExtension(string extension)
        {
            string value = (extension ?? string.Empty)
                .Trim()
                .TrimStart('.')
                .ToLowerInvariant();

            if (value == "png" || value == "jpg" || value == "jpeg" ||
                value == "gif" || value == "bmp")
            {
                return value;
            }

            throw new InvalidOperationException(
                "Unsupported image type for PPTX Writer: " + value +
                ". Supported types are PNG, JPEG, GIF and BMP.");
        }

        private static string ShapePreset(PresentationShapeKind kind)
        {
            if (kind == PresentationShapeKind.RoundedRectangle)
                return "roundRect";
            if (kind == PresentationShapeKind.Ellipse)
                return "ellipse";
            if (kind == PresentationShapeKind.Triangle)
                return "triangle";
            if (kind == PresentationShapeKind.Diamond)
                return "diamond";
            return "rect";
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

        private static void WriteExtendedProperties(ZipArchive archive, int slideCount)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>PowerPointLite</Application><PresentationFormat>Widescreen</PresentationFormat>" +
                "<Slides>" + slideCount.ToString() + "</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips>" +
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

        private static string AlignmentValue(PresentationTextAlignment alignment)
        {
            if (alignment == PresentationTextAlignment.Center)
                return "ctr";

            if (alignment == PresentationTextAlignment.Right)
                return "r";

            return "l";
        }

        private static string NormalizeColor(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "20242A";

            string candidate = value.Trim().TrimStart('#').ToUpperInvariant();

            if (candidate.Length != 6)
                return "20242A";

            for (int i = 0; i < candidate.Length; i++)
            {
                char ch = candidate[i];
                bool isHex =
                    (ch >= '0' && ch <= '9') ||
                    (ch >= 'A' && ch <= 'F');

                if (!isHex)
                    return "20242A";
            }

            return candidate;
        }

        private static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }
    }
}
