using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal static class AdvancedEditorThumbnailExtension
    {
        public static void Attach(
            AdvancedPresentationEditorForm editor,
            PresentationEditSession session)
        {
            if (editor == null || session == null)
                return;

            ListBox list = FindControl<ListBox>(editor);
            if (list == null)
                return;

            list.DrawMode = DrawMode.OwnerDrawFixed;
            list.ItemHeight = 92;
            list.DrawItem += delegate(object sender, DrawItemEventArgs e)
            {
                DrawSlideItem(list, session, e);
            };

            editor.Resize += delegate { list.Invalidate(); };
            editor.Shown += delegate { list.Invalidate(); };
        }

        private static void DrawSlideItem(
            ListBox list,
            PresentationEditSession session,
            DrawItemEventArgs e)
        {
            if (e.Index < 0 ||
                e.Index >= session.Document.Slides.Count)
            {
                return;
            }

            bool selected =
                (e.State & DrawItemState.Selected) ==
                DrawItemState.Selected;

            Color background = selected
                ? Color.FromArgb(54, 64, 82)
                : ApplicationTheme.Sidebar;

            using (Brush backgroundBrush = new SolidBrush(background))
                e.Graphics.FillRectangle(backgroundBrush, e.Bounds);

            Rectangle preview = new Rectangle(
                e.Bounds.Left + 8,
                e.Bounds.Top + 7,
                112,
                63);

            using (Bitmap thumbnail = RenderSlide(
                session.Document,
                e.Index,
                preview.Width * 2,
                preview.Height * 2))
            {
                e.Graphics.InterpolationMode =
                    InterpolationMode.HighQualityBicubic;
                e.Graphics.DrawImage(
                    thumbnail,
                    preview);
            }

            using (Pen border = new Pen(
                selected
                    ? ApplicationTheme.Accent
                    : Color.FromArgb(82, 88, 98),
                selected ? 2f : 1f))
            {
                e.Graphics.DrawRectangle(border, preview);
            }

            PresentationSlide slide =
                session.Document.Slides[e.Index];
            string title = GetSlideTitle(slide, e.Index);

            Rectangle textRect = new Rectangle(
                preview.Right + 9,
                e.Bounds.Top + 10,
                Math.Max(20, e.Bounds.Right - preview.Right - 14),
                56);

            using (Font numberFont = new Font("Segoe UI", 8f, FontStyle.Bold))
            using (Font titleFont = new Font("Segoe UI", 8.5f, FontStyle.Regular))
            using (Brush primary = new SolidBrush(ApplicationTheme.PrimaryText))
            using (Brush secondary = new SolidBrush(ApplicationTheme.SecondaryText))
            {
                e.Graphics.DrawString(
                    (e.Index + 1).ToString(),
                    numberFont,
                    secondary,
                    textRect.Left,
                    textRect.Top);

                RectangleF titleBounds = new RectangleF(
                    textRect.Left,
                    textRect.Top + 20,
                    textRect.Width,
                    textRect.Height - 20);

                using (StringFormat format = new StringFormat())
                {
                    format.Trimming = StringTrimming.EllipsisCharacter;
                    format.FormatFlags = StringFormatFlags.LineLimit;
                    e.Graphics.DrawString(
                        title,
                        titleFont,
                        primary,
                        titleBounds,
                        format);
                }
            }

            e.DrawFocusRectangle();
        }

        private static Bitmap RenderSlide(
            PresentationDocument document,
            int slideIndex,
            int width,
            int height)
        {
            Bitmap bitmap = new Bitmap(
                Math.Max(32, width),
                Math.Max(18, height),
                PixelFormat.Format32bppArgb);

            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;

                if (document == null ||
                    slideIndex < 0 ||
                    slideIndex >= document.Slides.Count)
                {
                    return bitmap;
                }

                PresentationSlide slide = document.Slides[slideIndex];
                if (slide == null)
                    return bitmap;

                for (int i = 0; i < slide.Shapes.Count; i++)
                    DrawShape(graphics, document, bitmap.Size, slide.Shapes[i]);

                for (int i = 0; i < slide.Images.Count; i++)
                    DrawImage(graphics, document, bitmap.Size, slide.Images[i]);

                for (int i = 0; i < slide.TextBoxes.Count; i++)
                    DrawText(graphics, document, bitmap.Size, slide.TextBoxes[i]);

                for (int i = 0; i < slide.Tables.Count; i++)
                    DrawTable(graphics, document, bitmap.Size, slide.Tables[i]);
            }

            return bitmap;
        }

        private static void DrawShape(
            Graphics graphics,
            PresentationDocument document,
            Size size,
            PresentationShape shape)
        {
            if (shape == null)
                return;

            RectangleF rect = ToPixels(
                document,
                size,
                shape.X,
                shape.Y,
                shape.Width,
                shape.Height);

            using (Brush fill = new SolidBrush(ParseColor(
                shape.FillColorHex,
                Color.FromArgb(91, 140, 255))))
            using (Pen line = new Pen(ParseColor(
                shape.LineColorHex,
                Color.FromArgb(53, 106, 230)),
                1f))
            {
                if (shape.Kind == PresentationShapeKind.Ellipse)
                {
                    graphics.FillEllipse(fill, rect);
                    graphics.DrawEllipse(line, rect);
                }
                else if (shape.Kind == PresentationShapeKind.Triangle)
                {
                    PointF[] points = new PointF[]
                    {
                        new PointF(rect.Left + rect.Width / 2f, rect.Top),
                        new PointF(rect.Right, rect.Bottom),
                        new PointF(rect.Left, rect.Bottom)
                    };
                    graphics.FillPolygon(fill, points);
                    graphics.DrawPolygon(line, points);
                }
                else if (shape.Kind == PresentationShapeKind.Diamond)
                {
                    PointF[] points = new PointF[]
                    {
                        new PointF(rect.Left + rect.Width / 2f, rect.Top),
                        new PointF(rect.Right, rect.Top + rect.Height / 2f),
                        new PointF(rect.Left + rect.Width / 2f, rect.Bottom),
                        new PointF(rect.Left, rect.Top + rect.Height / 2f)
                    };
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
            PresentationDocument document,
            Size size,
            PresentationImage item)
        {
            if (item == null || item.Data == null || item.Data.Length == 0)
                return;

            RectangleF rect = ToPixels(
                document,
                size,
                item.X,
                item.Y,
                item.Width,
                item.Height);

            try
            {
                using (MemoryStream stream = new MemoryStream(item.Data, false))
                using (Image image = Image.FromStream(stream))
                {
                    graphics.DrawImage(
                        image,
                        Rectangle.Round(rect));
                }
            }
            catch
            {
                using (Brush placeholder = new SolidBrush(Color.FromArgb(230, 232, 236)))
                    graphics.FillRectangle(placeholder, rect);
            }
        }

        private static void DrawText(
            Graphics graphics,
            PresentationDocument document,
            Size size,
            PresentationTextBox box)
        {
            if (box == null)
                return;

            RectangleF rect = ToPixels(
                document,
                size,
                box.X,
                box.Y,
                box.Width,
                box.Height);

            FontStyle style = FontStyle.Regular;
            if (box.Bold)
                style |= FontStyle.Bold;
            if (box.Italic)
                style |= FontStyle.Italic;

            float scale = size.Width / 960f;
            float fontSize = Math.Max(3f, box.FontSizePoints * scale);

            using (Font font = SafeFont(box.FontFamily, fontSize, style))
            using (Brush brush = new SolidBrush(ParseColor(
                box.ColorHex,
                Color.FromArgb(32, 36, 42))))
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
                graphics.DrawString(
                    box.Text ?? string.Empty,
                    font,
                    brush,
                    rect,
                    format);
            }
        }

        private static void DrawTable(
            Graphics graphics,
            PresentationDocument document,
            Size size,
            PresentationTable table)
        {
            if (table == null)
                return;

            RectangleF rect = ToPixels(
                document,
                size,
                table.X,
                table.Y,
                table.Width,
                table.Height);
            int rows = Math.Max(1, table.Rows);
            int columns = Math.Max(1, table.Columns);
            float cellWidth = rect.Width / columns;
            float cellHeight = rect.Height / rows;

            using (Pen border = new Pen(Color.FromArgb(172, 176, 184), 0.7f))
            {
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        PresentationTableCell cell = table.GetCell(row, column);
                        RectangleF cellRect = new RectangleF(
                            rect.Left + column * cellWidth,
                            rect.Top + row * cellHeight,
                            cellWidth,
                            cellHeight);

                        using (Brush fill = new SolidBrush(ParseColor(
                            cell == null ? null : cell.FillColorHex,
                            Color.White)))
                        {
                            graphics.FillRectangle(fill, cellRect);
                        }

                        graphics.DrawRectangle(
                            border,
                            Rectangle.Round(cellRect));
                    }
                }
            }
        }

        private static RectangleF ToPixels(
            PresentationDocument document,
            Size size,
            long x,
            long y,
            long width,
            long height)
        {
            long documentWidth = document.WidthEmu > 0
                ? document.WidthEmu
                : PresentationDocument.DefaultWidthEmu;
            long documentHeight = document.HeightEmu > 0
                ? document.HeightEmu
                : PresentationDocument.DefaultHeightEmu;

            return new RectangleF(
                (float)(x / (double)documentWidth * size.Width),
                (float)(y / (double)documentHeight * size.Height),
                Math.Max(1f, (float)(width / (double)documentWidth * size.Width)),
                Math.Max(1f, (float)(height / (double)documentHeight * size.Height)));
        }

        private static string GetSlideTitle(
            PresentationSlide slide,
            int index)
        {
            if (slide != null)
            {
                for (int i = 0; i < slide.TextBoxes.Count; i++)
                {
                    PresentationTextBox box = slide.TextBoxes[i];
                    if (box != null &&
                        !string.IsNullOrEmpty(box.Text) &&
                        string.Equals(
                            box.Name,
                            "Title",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return FirstLine(box.Text);
                    }
                }

                if (!string.IsNullOrEmpty(slide.Name))
                    return slide.Name;
            }

            return "Slide " + (index + 1).ToString();
        }

        private static string FirstLine(string value)
        {
            string text = value ?? string.Empty;
            int newline = text.IndexOfAny(new char[] { '\r', '\n' });
            if (newline >= 0)
                text = text.Substring(0, newline);
            return text.Length > 60 ? text.Substring(0, 60) + "…" : text;
        }

        private static Font SafeFont(
            string family,
            float size,
            FontStyle style)
        {
            try
            {
                return new Font(
                    string.IsNullOrEmpty(family) ? "Arial" : family,
                    size,
                    style,
                    GraphicsUnit.Point);
            }
            catch
            {
                return new Font(
                    SystemFonts.MessageBoxFont.FontFamily,
                    size,
                    style);
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

        private static T FindControl<T>(Control root)
            where T : Control
        {
            if (root == null)
                return null;

            T direct = root as T;
            if (direct != null)
                return direct;

            for (int i = 0; i < root.Controls.Count; i++)
            {
                T found = FindControl<T>(root.Controls[i]);
                if (found != null)
                    return found;
            }

            return null;
        }
    }
}
