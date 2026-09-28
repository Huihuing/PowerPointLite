using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;

namespace PptxViewer
{
    internal sealed class PdfRasterPage : IDisposable
    {
        public Bitmap Bitmap { get; private set; }
        public float WidthPoints { get; private set; }
        public float HeightPoints { get; private set; }

        public PdfRasterPage(Bitmap bitmap, float widthPoints, float heightPoints)
        {
            if (bitmap == null)
                throw new ArgumentNullException("bitmap");

            Bitmap = bitmap;
            WidthPoints = Math.Max(1f, widthPoints);
            HeightPoints = Math.Max(1f, heightPoints);
        }

        public void Dispose()
        {
            if (Bitmap != null)
            {
                Bitmap.Dispose();
                Bitmap = null;
            }
        }
    }

    internal static class RasterPdfWriter
    {
        public static void Write(string path, IList<PdfRasterPage> pages)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("PDF output path is required.", "path");
            if (pages == null || pages.Count == 0)
                throw new InvalidOperationException("A PDF must contain at least one page.");

            string fullPath = Path.GetFullPath(path);
            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string stage = fullPath + ".writing";
            if (File.Exists(stage))
                File.Delete(stage);

            try
            {
                using (FileStream output = new FileStream(
                    stage,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None))
                {
                    WriteDocument(output, pages);
                }

                ReplaceSafely(stage, fullPath);
            }
            catch
            {
                try { if (File.Exists(stage)) File.Delete(stage); }
                catch { }
                throw;
            }
        }

        private static void WriteDocument(Stream output, IList<PdfRasterPage> pages)
        {
            int objectCount = 2 + pages.Count * 3;
            long[] offsets = new long[objectCount + 1];

            WriteAscii(output, "%PDF-1.4\n%PPLT\n");

            WriteObject(output, offsets, 1,
                "<< /Type /Catalog /Pages 2 0 R >>");

            StringBuilder kids = new StringBuilder();
            for (int i = 0; i < pages.Count; i++)
            {
                if (i > 0)
                    kids.Append(' ');
                kids.Append((3 + i * 3).ToString(CultureInfo.InvariantCulture));
                kids.Append(" 0 R");
            }

            WriteObject(
                output,
                offsets,
                2,
                "<< /Type /Pages /Count " +
                pages.Count.ToString(CultureInfo.InvariantCulture) +
                " /Kids [" + kids.ToString() + "] >>");

            for (int i = 0; i < pages.Count; i++)
            {
                PdfRasterPage page = pages[i];
                if (page == null || page.Bitmap == null)
                    throw new InvalidOperationException("PDF page bitmap is missing.");

                int pageObject = 3 + i * 3;
                int contentObject = pageObject + 1;
                int imageObject = pageObject + 2;

                string width = page.WidthPoints.ToString("0.###", CultureInfo.InvariantCulture);
                string height = page.HeightPoints.ToString("0.###", CultureInfo.InvariantCulture);

                WriteObject(
                    output,
                    offsets,
                    pageObject,
                    "<< /Type /Page /Parent 2 0 R " +
                    "/MediaBox [0 0 " + width + " " + height + "] " +
                    "/Resources << /XObject << /Im0 " + imageObject.ToString() +
                    " 0 R >> >> /Contents " + contentObject.ToString() + " 0 R >>");

                string content =
                    "q\n" + width + " 0 0 " + height + " 0 0 cm\n" +
                    "/Im0 Do\nQ\n";
                byte[] contentBytes = Encoding.ASCII.GetBytes(content);

                offsets[contentObject] = output.Position;
                WriteAscii(output, contentObject.ToString() + " 0 obj\n");
                WriteAscii(output, "<< /Length " + contentBytes.Length.ToString() + " >>\nstream\n");
                output.Write(contentBytes, 0, contentBytes.Length);
                WriteAscii(output, "endstream\nendobj\n");

                byte[] jpeg = EncodeJpeg(page.Bitmap, 90L);

                offsets[imageObject] = output.Position;
                WriteAscii(output, imageObject.ToString() + " 0 obj\n");
                WriteAscii(
                    output,
                    "<< /Type /XObject /Subtype /Image " +
                    "/Width " + page.Bitmap.Width.ToString() + " " +
                    "/Height " + page.Bitmap.Height.ToString() + " " +
                    "/ColorSpace /DeviceRGB /BitsPerComponent 8 " +
                    "/Filter /DCTDecode /Length " + jpeg.Length.ToString() + " >>\nstream\n");
                output.Write(jpeg, 0, jpeg.Length);
                WriteAscii(output, "\nendstream\nendobj\n");
            }

            long xref = output.Position;
            WriteAscii(output, "xref\n0 " + (objectCount + 1).ToString() + "\n");
            WriteAscii(output, "0000000000 65535 f \n");

            for (int i = 1; i <= objectCount; i++)
            {
                WriteAscii(
                    output,
                    offsets[i].ToString("0000000000", CultureInfo.InvariantCulture) +
                    " 00000 n \n");
            }

            WriteAscii(
                output,
                "trailer\n<< /Size " + (objectCount + 1).ToString() +
                " /Root 1 0 R >>\nstartxref\n" +
                xref.ToString(CultureInfo.InvariantCulture) +
                "\n%%EOF\n");
        }

        private static void WriteObject(
            Stream output,
            long[] offsets,
            int objectNumber,
            string body)
        {
            offsets[objectNumber] = output.Position;
            WriteAscii(output, objectNumber.ToString() + " 0 obj\n");
            WriteAscii(output, body);
            WriteAscii(output, "\nendobj\n");
        }

        private static byte[] EncodeJpeg(Bitmap source, long quality)
        {
            using (Bitmap flattened = new Bitmap(
                Math.Max(1, source.Width),
                Math.Max(1, source.Height),
                PixelFormat.Format24bppRgb))
            {
                using (Graphics graphics = Graphics.FromImage(flattened))
                {
                    graphics.Clear(Color.White);
                    graphics.DrawImage(
                        source,
                        new Rectangle(0, 0, flattened.Width, flattened.Height));
                }

                using (MemoryStream buffer = new MemoryStream())
                {
                    ImageCodecInfo jpegCodec = FindEncoder(ImageFormat.Jpeg);
                    if (jpegCodec == null)
                    {
                        flattened.Save(buffer, ImageFormat.Jpeg);
                    }
                    else
                    {
                        using (EncoderParameters parameters = new EncoderParameters(1))
                        {
                            parameters.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality,
                                Math.Max(1L, Math.Min(100L, quality)));
                            flattened.Save(buffer, jpegCodec, parameters);
                        }
                    }

                    return buffer.ToArray();
                }
            }
        }

        private static ImageCodecInfo FindEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            for (int i = 0; i < codecs.Length; i++)
            {
                if (codecs[i].FormatID == format.Guid)
                    return codecs[i];
            }
            return null;
        }

        private static void ReplaceSafely(string stage, string destination)
        {
            if (!File.Exists(destination))
            {
                File.Move(stage, destination);
                return;
            }

            string backup = destination + ".backup";
            if (File.Exists(backup))
                File.Delete(backup);

            File.Move(destination, backup);

            try
            {
                File.Move(stage, destination);
                File.Delete(backup);
            }
            catch
            {
                try
                {
                    if (File.Exists(destination))
                        File.Delete(destination);
                    if (File.Exists(backup))
                        File.Move(backup, destination);
                }
                catch { }
                throw;
            }
        }

        private static void WriteAscii(Stream output, string text)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(text ?? string.Empty);
            output.Write(bytes, 0, bytes.Length);
        }
    }

    internal static class PdfExportService
    {
        public static void ExportPresentation(
            PresentationDocument document,
            string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            List<PdfRasterPage> pages = new List<PdfRasterPage>();
            try
            {
                long widthEmu = document.WidthEmu > 0
                    ? document.WidthEmu
                    : PresentationDocument.DefaultWidthEmu;
                long heightEmu = document.HeightEmu > 0
                    ? document.HeightEmu
                    : PresentationDocument.DefaultHeightEmu;

                float pageWidth = Math.Max(72f, widthEmu / 12700f);
                float pageHeight = Math.Max(72f, heightEmu / 12700f);

                int bitmapWidth = 1600;
                int bitmapHeight = Math.Max(
                    1,
                    (int)Math.Round(bitmapWidth * pageHeight / pageWidth));

                for (int i = 0; i < document.Slides.Count; i++)
                {
                    Bitmap bitmap = new Bitmap(bitmapWidth, bitmapHeight);
                    RenderPresentationSlide(
                        bitmap,
                        document,
                        document.Slides[i]);
                    pages.Add(new PdfRasterPage(bitmap, pageWidth, pageHeight));
                }

                RasterPdfWriter.Write(outputPath, pages);
            }
            finally
            {
                DisposePages(pages);
            }
        }

        public static void ExportTextDocument(
            TextDocument document,
            string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            List<PdfRasterPage> pages = RenderTextDocumentPages(document);
            try
            {
                RasterPdfWriter.Write(outputPath, pages);
            }
            finally
            {
                DisposePages(pages);
            }
        }

        public static void ExportSpreadsheet(
            SpreadsheetDocument document,
            string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            List<PdfRasterPage> pages = RenderSpreadsheetPages(document);
            try
            {
                RasterPdfWriter.Write(outputPath, pages);
            }
            finally
            {
                DisposePages(pages);
            }
        }

        private static void RenderPresentationSlide(
            Bitmap bitmap,
            PresentationDocument document,
            PresentationSlide slide)
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                graphics.Clear(Color.White);

                if (slide == null)
                    return;

                long widthEmu = document.WidthEmu > 0
                    ? document.WidthEmu
                    : PresentationDocument.DefaultWidthEmu;
                long heightEmu = document.HeightEmu > 0
                    ? document.HeightEmu
                    : PresentationDocument.DefaultHeightEmu;

                for (int i = 0; i < slide.Shapes.Count; i++)
                    DrawShape(graphics, bitmap.Size, widthEmu, heightEmu, slide.Shapes[i]);

                for (int i = 0; i < slide.Images.Count; i++)
                    DrawImage(graphics, bitmap.Size, widthEmu, heightEmu, slide.Images[i]);

                for (int i = 0; i < slide.Tables.Count; i++)
                    DrawTable(graphics, bitmap.Size, widthEmu, heightEmu, slide.Tables[i]);

                for (int i = 0; i < slide.TextBoxes.Count; i++)
                    DrawTextBox(graphics, bitmap.Size, widthEmu, heightEmu, slide.TextBoxes[i]);
            }
        }

        private static void DrawTextBox(
            Graphics graphics,
            Size canvas,
            long documentWidth,
            long documentHeight,
            PresentationTextBox box)
        {
            if (box == null)
                return;

            RectangleF rect = MapRect(
                canvas,
                documentWidth,
                documentHeight,
                box.X,
                box.Y,
                box.Width,
                box.Height);

            FontStyle style = FontStyle.Regular;
            if (box.Bold) style |= FontStyle.Bold;
            if (box.Italic) style |= FontStyle.Italic;

            float points = Math.Max(1f, box.FontSizePoints);
            using (Font font = SafeFont(box.FontFamily, points, style))
            using (Brush brush = new SolidBrush(ParseColor(box.ColorHex, Color.FromArgb(32, 36, 42))))
            using (StringFormat format = new StringFormat())
            {
                format.Alignment =
                    box.Alignment == PresentationTextAlignment.Center
                        ? StringAlignment.Center
                        : box.Alignment == PresentationTextAlignment.Right
                            ? StringAlignment.Far
                            : StringAlignment.Near;
                format.LineAlignment = StringAlignment.Center;
                format.Trimming = StringTrimming.EllipsisCharacter;
                graphics.DrawString(box.Text ?? string.Empty, font, brush, rect, format);
            }
        }

        private static void DrawShape(
            Graphics graphics,
            Size canvas,
            long documentWidth,
            long documentHeight,
            PresentationShape shape)
        {
            if (shape == null)
                return;

            RectangleF rect = MapRect(
                canvas,
                documentWidth,
                documentHeight,
                shape.X,
                shape.Y,
                shape.Width,
                shape.Height);

            using (Brush fill = new SolidBrush(ParseColor(shape.FillColorHex, Color.FromArgb(91, 140, 255))))
            using (Pen line = new Pen(
                ParseColor(shape.LineColorHex, Color.FromArgb(53, 106, 230)),
                Math.Max(1f, shape.LineWidthPoints * 1.5f)))
            {
                if (shape.Kind == PresentationShapeKind.Ellipse)
                {
                    graphics.FillEllipse(fill, rect);
                    graphics.DrawEllipse(line, rect);
                }
                else if (shape.Kind == PresentationShapeKind.Triangle ||
                    shape.Kind == PresentationShapeKind.Diamond)
                {
                    PointF[] points;
                    if (shape.Kind == PresentationShapeKind.Triangle)
                    {
                        points = new PointF[]
                        {
                            new PointF(rect.Left + rect.Width / 2f, rect.Top),
                            new PointF(rect.Right, rect.Bottom),
                            new PointF(rect.Left, rect.Bottom)
                        };
                    }
                    else
                    {
                        points = new PointF[]
                        {
                            new PointF(rect.Left + rect.Width / 2f, rect.Top),
                            new PointF(rect.Right, rect.Top + rect.Height / 2f),
                            new PointF(rect.Left + rect.Width / 2f, rect.Bottom),
                            new PointF(rect.Left, rect.Top + rect.Height / 2f)
                        };
                    }

                    graphics.FillPolygon(fill, points);
                    graphics.DrawPolygon(line, points);
                }
                else
                {
                    graphics.FillRectangle(fill, rect);
                    graphics.DrawRectangle(line, Rectangle.Round(rect));
                }
            }
        }

        private static void DrawImage(
            Graphics graphics,
            Size canvas,
            long documentWidth,
            long documentHeight,
            PresentationImage image)
        {
            if (image == null || image.Data == null || image.Data.Length == 0)
                return;

            RectangleF rect = MapRect(
                canvas,
                documentWidth,
                documentHeight,
                image.X,
                image.Y,
                image.Width,
                image.Height);

            try
            {
                using (MemoryStream stream = new MemoryStream(image.Data, false))
                using (Image decoded = Image.FromStream(stream))
                {
                    graphics.DrawImage(decoded, Rectangle.Round(rect));
                }
            }
            catch
            {
                using (Pen border = new Pen(Color.FromArgb(180, 80, 80), 2f))
                    graphics.DrawRectangle(border, Rectangle.Round(rect));
            }
        }

        private static void DrawTable(
            Graphics graphics,
            Size canvas,
            long documentWidth,
            long documentHeight,
            PresentationTable table)
        {
            if (table == null || table.Rows <= 0 || table.Columns <= 0)
                return;

            RectangleF rect = MapRect(
                canvas,
                documentWidth,
                documentHeight,
                table.X,
                table.Y,
                table.Width,
                table.Height);

            float cellWidth = rect.Width / table.Columns;
            float cellHeight = rect.Height / table.Rows;

            for (int row = 0; row < table.Rows; row++)
            {
                for (int column = 0; column < table.Columns; column++)
                {
                    PresentationTableCell cell = table.GetCell(row, column);
                    RectangleF cellRect = new RectangleF(
                        rect.Left + column * cellWidth,
                        rect.Top + row * cellHeight,
                        cellWidth,
                        cellHeight);

                    Color fillColor = cell == null
                        ? Color.White
                        : ParseColor(cell.FillColorHex, Color.White);

                    using (Brush fill = new SolidBrush(fillColor))
                        graphics.FillRectangle(fill, cellRect);
                    using (Pen border = new Pen(Color.FromArgb(190, 194, 202), 1f))
                        graphics.DrawRectangle(border, Rectangle.Round(cellRect));

                    if (cell == null)
                        continue;

                    FontStyle style = cell.Bold ? FontStyle.Bold : FontStyle.Regular;
                    using (Font font = SafeFont(cell.FontFamily, cell.FontSizePoints, style))
                    using (Brush textBrush = new SolidBrush(ParseColor(cell.TextColorHex, Color.FromArgb(32, 36, 42))))
                    using (StringFormat format = new StringFormat())
                    {
                        format.LineAlignment = StringAlignment.Center;
                        format.Alignment =
                            cell.Alignment == PresentationTextAlignment.Center
                                ? StringAlignment.Center
                                : cell.Alignment == PresentationTextAlignment.Right
                                    ? StringAlignment.Far
                                    : StringAlignment.Near;
                        RectangleF padded = RectangleF.Inflate(cellRect, -5f, -3f);
                        graphics.DrawString(cell.Text ?? string.Empty, font, textBrush, padded, format);
                    }
                }
            }
        }

        private static List<PdfRasterPage> RenderTextDocumentPages(TextDocument document)
        {
            const int pixelWidth = 1240;
            const int pixelHeight = 1754;
            const float pageWidthPoints = 595.28f;
            const float pageHeightPoints = 841.89f;
            const float margin = 100f;

            List<PdfRasterPage> pages = new List<PdfRasterPage>();
            Bitmap bitmap = CreateWhiteBitmap(pixelWidth, pixelHeight);
            Graphics graphics = Graphics.FromImage(bitmap);
            ConfigureGraphics(graphics);
            float y = margin;

            try
            {
                if (!string.IsNullOrEmpty(document.Title))
                {
                    using (Font titleFont = SafeFont("Arial", 18f, FontStyle.Bold))
                    using (Brush brush = new SolidBrush(Color.FromArgb(32, 36, 42)))
                    {
                        graphics.DrawString(document.Title, titleFont, brush, margin, y);
                        y += titleFont.GetHeight(graphics) + 28f;
                    }
                }

                for (int i = 0; i < document.Paragraphs.Count; i++)
                {
                    DocumentParagraph paragraph = document.Paragraphs[i];
                    string text = paragraph == null ? string.Empty : paragraph.PlainText;
                    DocumentTextRun sample =
                        paragraph != null && paragraph.Runs.Count > 0
                            ? paragraph.Runs[0]
                            : null;

                    string family = sample == null ? "Arial" : sample.FontFamily;
                    float fontSize = sample == null ? 11f : sample.FontSizePoints;
                    FontStyle style = FontStyle.Regular;
                    if (sample != null && sample.Bold) style |= FontStyle.Bold;
                    if (sample != null && sample.Italic) style |= FontStyle.Italic;
                    if (sample != null && sample.Underline) style |= FontStyle.Underline;
                    Color color = sample == null
                        ? Color.FromArgb(32, 36, 42)
                        : ParseColor(sample.ColorHex, Color.FromArgb(32, 36, 42));

                    using (Font font = SafeFont(family, Math.Max(1f, fontSize), style))
                    using (Brush brush = new SolidBrush(color))
                    using (StringFormat format = new StringFormat())
                    {
                        format.Alignment = ParagraphAlignment(paragraph);
                        format.Trimming = StringTrimming.Word;

                        float availableWidth = pixelWidth - margin * 2f;
                        SizeF measured = graphics.MeasureString(
                            string.IsNullOrEmpty(text) ? " " : text,
                            font,
                            new SizeF(availableWidth, pixelHeight),
                            format);
                        float blockHeight = Math.Max(font.GetHeight(graphics), measured.Height) + 12f;

                        if (y + blockHeight > pixelHeight - margin && y > margin)
                        {
                            graphics.Dispose();
                            pages.Add(new PdfRasterPage(bitmap, pageWidthPoints, pageHeightPoints));
                            bitmap = CreateWhiteBitmap(pixelWidth, pixelHeight);
                            graphics = Graphics.FromImage(bitmap);
                            ConfigureGraphics(graphics);
                            y = margin;
                        }

                        RectangleF layout = new RectangleF(
                            margin,
                            y,
                            availableWidth,
                            Math.Max(font.GetHeight(graphics), measured.Height + 4f));
                        graphics.DrawString(text ?? string.Empty, font, brush, layout, format);
                        y += blockHeight;
                    }
                }
            }
            finally
            {
                graphics.Dispose();
            }

            pages.Add(new PdfRasterPage(bitmap, pageWidthPoints, pageHeightPoints));
            return pages;
        }

        private static List<PdfRasterPage> RenderSpreadsheetPages(SpreadsheetDocument document)
        {
            const int pixelWidth = 1240;
            const int pixelHeight = 1754;
            const float pageWidthPoints = 595.28f;
            const float pageHeightPoints = 841.89f;
            const int rowsPerPage = 34;
            const int columnsPerPage = 8;
            const float margin = 70f;
            const float headerHeight = 54f;

            List<PdfRasterPage> pages = new List<PdfRasterPage>();

            for (int sheetIndex = 0; sheetIndex < document.Sheets.Count; sheetIndex++)
            {
                SpreadsheetSheet sheet = document.Sheets[sheetIndex];
                if (sheet == null)
                    continue;

                int maxRow = Math.Max(1, sheet.MaxRow);
                int maxColumn = Math.Max(1, sheet.MaxColumn);

                for (int rowStart = 1; rowStart <= maxRow; rowStart += rowsPerPage)
                {
                    for (int columnStart = 1; columnStart <= maxColumn; columnStart += columnsPerPage)
                    {
                        Bitmap bitmap = CreateWhiteBitmap(pixelWidth, pixelHeight);
                        using (Graphics graphics = Graphics.FromImage(bitmap))
                        {
                            ConfigureGraphics(graphics);

                            using (Font titleFont = SafeFont("Arial", 14f, FontStyle.Bold))
                            using (Brush titleBrush = new SolidBrush(Color.FromArgb(32, 36, 42)))
                            {
                                string title =
                                    (string.IsNullOrEmpty(sheet.Name) ? "Sheet" : sheet.Name) +
                                    "  •  rows " + rowStart.ToString() + "-" +
                                    Math.Min(maxRow, rowStart + rowsPerPage - 1).ToString();
                                graphics.DrawString(title, titleFont, titleBrush, margin, 30f);
                            }

                            float gridTop = margin + headerHeight;
                            float gridWidth = pixelWidth - margin * 2f;
                            float gridHeight = pixelHeight - gridTop - margin;
                            float cellWidth = gridWidth / columnsPerPage;
                            float cellHeight = gridHeight / rowsPerPage;

                            using (Font cellFont = SafeFont("Arial", 8.5f, FontStyle.Regular))
                            using (Pen gridPen = new Pen(Color.FromArgb(210, 214, 220), 1f))
                            using (Brush textBrush = new SolidBrush(Color.FromArgb(32, 36, 42)))
                            {
                                for (int r = 0; r < rowsPerPage; r++)
                                {
                                    int row = rowStart + r;
                                    if (row > maxRow)
                                        break;

                                    for (int c = 0; c < columnsPerPage; c++)
                                    {
                                        int column = columnStart + c;
                                        if (column > maxColumn)
                                            break;

                                        RectangleF rect = new RectangleF(
                                            margin + c * cellWidth,
                                            gridTop + r * cellHeight,
                                            cellWidth,
                                            cellHeight);
                                        graphics.DrawRectangle(gridPen, Rectangle.Round(rect));

                                        SpreadsheetCell cell = sheet.GetCell(row, column, false);
                                        if (cell == null)
                                            continue;

                                        RectangleF textRect = RectangleF.Inflate(rect, -4f, -2f);
                                        using (StringFormat format = new StringFormat())
                                        {
                                            format.Trimming = StringTrimming.EllipsisCharacter;
                                            format.LineAlignment = StringAlignment.Center;
                                            graphics.DrawString(
                                                cell.DisplayText,
                                                cellFont,
                                                textBrush,
                                                textRect,
                                                format);
                                        }
                                    }
                                }
                            }
                        }

                        pages.Add(new PdfRasterPage(bitmap, pageWidthPoints, pageHeightPoints));
                    }
                }
            }

            if (pages.Count == 0)
            {
                pages.Add(new PdfRasterPage(
                    CreateWhiteBitmap(pixelWidth, pixelHeight),
                    pageWidthPoints,
                    pageHeightPoints));
            }

            return pages;
        }

        private static RectangleF MapRect(
            Size canvas,
            long documentWidth,
            long documentHeight,
            long x,
            long y,
            long width,
            long height)
        {
            documentWidth = Math.Max(1L, documentWidth);
            documentHeight = Math.Max(1L, documentHeight);

            return new RectangleF(
                (float)(Math.Max(0L, x) / (double)documentWidth * canvas.Width),
                (float)(Math.Max(0L, y) / (double)documentHeight * canvas.Height),
                (float)(Math.Max(1L, width) / (double)documentWidth * canvas.Width),
                (float)(Math.Max(1L, height) / (double)documentHeight * canvas.Height));
        }

        private static StringAlignment ParagraphAlignment(DocumentParagraph paragraph)
        {
            if (paragraph == null)
                return StringAlignment.Near;
            if (paragraph.Alignment == DocumentParagraphAlignment.Center)
                return StringAlignment.Center;
            if (paragraph.Alignment == DocumentParagraphAlignment.Right)
                return StringAlignment.Far;
            return StringAlignment.Near;
        }

        private static Bitmap CreateWhiteBitmap(int width, int height)
        {
            Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (Graphics graphics = Graphics.FromImage(bitmap))
                graphics.Clear(Color.White);
            return bitmap;
        }

        private static void ConfigureGraphics(Graphics graphics)
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
        }

        private static Font SafeFont(string family, float size, FontStyle style)
        {
            try
            {
                return new Font(
                    string.IsNullOrEmpty(family) ? "Arial" : family,
                    Math.Max(1f, size),
                    style,
                    GraphicsUnit.Point);
            }
            catch
            {
                return new Font(SystemFonts.MessageBoxFont.FontFamily, Math.Max(1f, size), style);
            }
        }

        private static Color ParseColor(string value, Color fallback)
        {
            string candidate = (value ?? string.Empty)
                .Trim()
                .TrimStart('#');

            int parsed;
            if (candidate.Length == 6 &&
                int.TryParse(
                    candidate,
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return Color.FromArgb(
                    (parsed >> 16) & 0xFF,
                    (parsed >> 8) & 0xFF,
                    parsed & 0xFF);
            }

            return fallback;
        }

        private static void DisposePages(IList<PdfRasterPage> pages)
        {
            if (pages == null)
                return;
            for (int i = 0; i < pages.Count; i++)
            {
                if (pages[i] != null)
                    pages[i].Dispose();
            }
        }
    }
}
