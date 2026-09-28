using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class PresentationEditorForm : Form
    {
        private readonly PresentationEditSession session;
        private readonly ListBox slideList;
        private readonly PresentationCanvas canvas;
        private readonly TextBox textEditor;
        private readonly ComboBox fontPicker;
        private readonly NumericUpDown fontSize;
        private readonly CheckBox boldCheck;
        private readonly CheckBox italicCheck;
        private readonly ComboBox alignmentPicker;
        private readonly TextBox colorEditor;
        private readonly Label statusLabel;

        private bool loadingProperties;

        public string SavedFilePath
        {
            get { return session.FilePath; }
        }

        public PresentationEditorForm(PresentationEditSession editSession)
        {
            if (editSession == null)
                throw new ArgumentNullException("editSession");

            session = editSession;

            Text = "New Presentation - Editor";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1280;
            Height = 800;
            MinimumSize = new Size(960, 640);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);

            Panel toolbar = new Panel();
            toolbar.Dock = DockStyle.Top;
            toolbar.Height = 52;
            toolbar.Padding = new Padding(8, 10, 8, 10);
            toolbar.BackColor = ApplicationTheme.Toolbar;
            Controls.Add(toolbar);

            Button addSlide = MakeToolbarButton("+ Slide", 8, 76);
            Button deleteSlide = MakeToolbarButton("Delete", 90, 72);
            Button moveUp = MakeToolbarButton("Up", 168, 56);
            Button moveDown = MakeToolbarButton("Down", 230, 64);
            Button save = MakeToolbarButton("Save", 306, 68);
            Button saveAs = MakeToolbarButton("Save As", 380, 78);

            toolbar.Controls.Add(addSlide);
            toolbar.Controls.Add(deleteSlide);
            toolbar.Controls.Add(moveUp);
            toolbar.Controls.Add(moveDown);
            toolbar.Controls.Add(save);
            toolbar.Controls.Add(saveAs);

            addSlide.Click += delegate { AddSlide(); };
            deleteSlide.Click += delegate { DeleteSlide(); };
            moveUp.Click += delegate { MoveSelectedSlide(-1); };
            moveDown.Click += delegate { MoveSelectedSlide(1); };
            save.Click += delegate { SaveDocument(false); };
            saveAs.Click += delegate { SaveDocument(true); };

            statusLabel = new Label();
            statusLabel.Dock = DockStyle.Bottom;
            statusLabel.Height = 28;
            statusLabel.Padding = new Padding(10, 0, 8, 0);
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusLabel.BackColor = ApplicationTheme.Toolbar;
            statusLabel.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(statusLabel);

            Panel propertyPanel = new Panel();
            propertyPanel.Dock = DockStyle.Right;
            propertyPanel.Width = 300;
            propertyPanel.Padding = new Padding(14);
            propertyPanel.BackColor = ApplicationTheme.Sidebar;
            Controls.Add(propertyPanel);

            Label propertyTitle = MakeLabel("Text properties", 0, 0, 250, 28, true);
            propertyPanel.Controls.Add(propertyTitle);

            Label textLabel = MakeLabel("Text", 0, 38, 250, 20, false);
            propertyPanel.Controls.Add(textLabel);

            textEditor = new TextBox();
            textEditor.Left = 0;
            textEditor.Top = 62;
            textEditor.Width = 270;
            textEditor.Height = 160;
            textEditor.Multiline = true;
            textEditor.ScrollBars = ScrollBars.Vertical;
            textEditor.BackColor = Color.FromArgb(245, 246, 248);
            textEditor.ForeColor = Color.FromArgb(28, 31, 36);
            textEditor.BorderStyle = BorderStyle.FixedSingle;
            propertyPanel.Controls.Add(textEditor);

            Label fontLabel = MakeLabel("Font", 0, 234, 250, 20, false);
            propertyPanel.Controls.Add(fontLabel);

            fontPicker = new ComboBox();
            fontPicker.Left = 0;
            fontPicker.Top = 258;
            fontPicker.Width = 270;
            fontPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            fontPicker.BackColor = Color.White;
            fontPicker.ForeColor = Color.FromArgb(28, 31, 36);
            propertyPanel.Controls.Add(fontPicker);

            List<string> installedFonts = SystemFontCatalog.GetInstalledFamilyNames();
            for (int i = 0; i < installedFonts.Count; i++)
                fontPicker.Items.Add(installedFonts[i]);

            Label sizeLabel = MakeLabel("Size", 0, 292, 70, 20, false);
            propertyPanel.Controls.Add(sizeLabel);

            fontSize = new NumericUpDown();
            fontSize.Left = 0;
            fontSize.Top = 316;
            fontSize.Width = 82;
            fontSize.Minimum = 1;
            fontSize.Maximum = 400;
            fontSize.DecimalPlaces = 1;
            fontSize.Increment = 0.5M;
            propertyPanel.Controls.Add(fontSize);

            boldCheck = new CheckBox();
            boldCheck.Text = "Bold";
            boldCheck.Left = 96;
            boldCheck.Top = 318;
            boldCheck.Width = 68;
            boldCheck.ForeColor = ApplicationTheme.PrimaryText;
            propertyPanel.Controls.Add(boldCheck);

            italicCheck = new CheckBox();
            italicCheck.Text = "Italic";
            italicCheck.Left = 174;
            italicCheck.Top = 318;
            italicCheck.Width = 70;
            italicCheck.ForeColor = ApplicationTheme.PrimaryText;
            propertyPanel.Controls.Add(italicCheck);

            Label alignLabel = MakeLabel("Alignment", 0, 354, 100, 20, false);
            propertyPanel.Controls.Add(alignLabel);

            alignmentPicker = new ComboBox();
            alignmentPicker.Left = 0;
            alignmentPicker.Top = 378;
            alignmentPicker.Width = 132;
            alignmentPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            alignmentPicker.Items.Add("Left");
            alignmentPicker.Items.Add("Center");
            alignmentPicker.Items.Add("Right");
            propertyPanel.Controls.Add(alignmentPicker);

            Label colorLabel = MakeLabel("Text color (#RRGGBB)", 0, 414, 180, 20, false);
            propertyPanel.Controls.Add(colorLabel);

            colorEditor = new TextBox();
            colorEditor.Left = 0;
            colorEditor.Top = 438;
            colorEditor.Width = 132;
            colorEditor.MaxLength = 7;
            propertyPanel.Controls.Add(colorEditor);

            Panel slidePanel = new Panel();
            slidePanel.Dock = DockStyle.Left;
            slidePanel.Width = 220;
            slidePanel.Padding = new Padding(10);
            slidePanel.BackColor = ApplicationTheme.Sidebar;
            Controls.Add(slidePanel);

            Label slidesTitle = new Label();
            slidesTitle.Dock = DockStyle.Top;
            slidesTitle.Height = 28;
            slidesTitle.Text = "Slides";
            slidesTitle.Font = new Font(Font, FontStyle.Bold);
            slidesTitle.ForeColor = ApplicationTheme.PrimaryText;
            slidePanel.Controls.Add(slidesTitle);

            slideList = new ListBox();
            slideList.Dock = DockStyle.Fill;
            slideList.BorderStyle = BorderStyle.None;
            slideList.BackColor = ApplicationTheme.Sidebar;
            slideList.ForeColor = ApplicationTheme.PrimaryText;
            slideList.IntegralHeight = false;
            slidePanel.Controls.Add(slideList);
            slideList.BringToFront();

            Panel canvasHost = new Panel();
            canvasHost.Dock = DockStyle.Fill;
            canvasHost.Padding = new Padding(24);
            canvasHost.BackColor = ApplicationTheme.Canvas;
            Controls.Add(canvasHost);
            canvasHost.BringToFront();

            canvas = new PresentationCanvas();
            canvas.Dock = DockStyle.Fill;
            canvas.Document = session.Document;
            canvas.BackColor = ApplicationTheme.Canvas;
            canvasHost.Controls.Add(canvas);

            slideList.SelectedIndexChanged += delegate
            {
                canvas.SelectedSlideIndex = slideList.SelectedIndex;
                canvas.SelectedTextBoxIndex = -1;
                canvas.Invalidate();
                LoadSelectedProperties();
            };

            canvas.SelectionChanged += delegate
            {
                LoadSelectedProperties();
            };

            textEditor.TextChanged += delegate { ApplyPropertyChanges(); };
            fontPicker.SelectedIndexChanged += delegate { ApplyPropertyChanges(); };
            fontSize.ValueChanged += delegate { ApplyPropertyChanges(); };
            boldCheck.CheckedChanged += delegate { ApplyPropertyChanges(); };
            italicCheck.CheckedChanged += delegate { ApplyPropertyChanges(); };
            alignmentPicker.SelectedIndexChanged += delegate { ApplyPropertyChanges(); };
            colorEditor.TextChanged += delegate { ApplyPropertyChanges(); };

            FormClosing += OnEditorClosing;

            RefreshSlideList();
            if (slideList.Items.Count > 0)
                slideList.SelectedIndex = 0;
            UpdateStatus();
        }

        public static PresentationEditorForm CreateNew(string title)
        {
            return new PresentationEditorForm(
                PresentationEditSession.CreateNew(title));
        }

        private Button MakeToolbarButton(string text, int left, int width)
        {
            Button button = new Button();
            button.Text = text;
            button.Left = left;
            button.Top = 10;
            button.Width = width;
            button.Height = 32;
            ApplicationTheme.ApplyButton(button);
            return button;
        }

        private Label MakeLabel(
            string text,
            int left,
            int top,
            int width,
            int height,
            bool bold)
        {
            Label label = new Label();
            label.Text = text;
            label.Left = left;
            label.Top = top;
            label.Width = width;
            label.Height = height;
            label.ForeColor = ApplicationTheme.PrimaryText;

            if (bold)
                label.Font = new Font(Font, FontStyle.Bold);

            return label;
        }

        private void AddSlide()
        {
            PresentationSlide slide = session.AddSlide("New Slide");
            slide.Name = "Slide " + session.Document.Slides.Count.ToString();
            RefreshSlideList();
            slideList.SelectedIndex = session.Document.Slides.Count - 1;
            UpdateStatus();
        }

        private void DeleteSlide()
        {
            int index = slideList.SelectedIndex;
            if (index < 0)
                return;

            if (!session.RemoveSlide(index))
            {
                MessageBox.Show(
                    this,
                    "A presentation must keep at least one slide.",
                    "Delete slide",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            RefreshSlideList();
            slideList.SelectedIndex = Math.Min(index, slideList.Items.Count - 1);
            UpdateStatus();
        }

        private void MoveSelectedSlide(int direction)
        {
            int from = slideList.SelectedIndex;
            int to = from + direction;

            if (!session.MoveSlide(from, to))
                return;

            RefreshSlideList();
            slideList.SelectedIndex = to;
            UpdateStatus();
        }

        private void RefreshSlideList()
        {
            int selected = slideList.SelectedIndex;
            slideList.BeginUpdate();
            slideList.Items.Clear();

            for (int i = 0; i < session.Document.Slides.Count; i++)
            {
                PresentationSlide slide = session.Document.Slides[i];
                string name = slide == null || string.IsNullOrEmpty(slide.Name)
                    ? "Slide " + (i + 1).ToString()
                    : slide.Name;
                slideList.Items.Add((i + 1).ToString() + ".  " + name);
            }

            slideList.EndUpdate();

            if (slideList.Items.Count > 0 && selected >= 0)
                slideList.SelectedIndex = Math.Min(selected, slideList.Items.Count - 1);

            canvas.Document = session.Document;
            canvas.Invalidate();
        }

        private PresentationTextBox GetSelectedTextBox()
        {
            int slideIndex = canvas.SelectedSlideIndex;
            int boxIndex = canvas.SelectedTextBoxIndex;

            if (slideIndex < 0 || slideIndex >= session.Document.Slides.Count)
                return null;

            PresentationSlide slide = session.Document.Slides[slideIndex];
            if (slide == null || boxIndex < 0 || boxIndex >= slide.TextBoxes.Count)
                return null;

            return slide.TextBoxes[boxIndex];
        }

        private void LoadSelectedProperties()
        {
            loadingProperties = true;

            try
            {
                PresentationTextBox box = GetSelectedTextBox();
                bool enabled = box != null;

                textEditor.Enabled = enabled;
                fontPicker.Enabled = enabled;
                fontSize.Enabled = enabled;
                boldCheck.Enabled = enabled;
                italicCheck.Enabled = enabled;
                alignmentPicker.Enabled = enabled;
                colorEditor.Enabled = enabled;

                if (!enabled)
                {
                    textEditor.Text = string.Empty;
                    colorEditor.Text = string.Empty;
                    return;
                }

                textEditor.Text = box.Text ?? string.Empty;
                SelectFontFamily(box.FontFamily);
                decimal size = (decimal)Math.Max(1f, Math.Min(400f, box.FontSizePoints));
                fontSize.Value = size;
                boldCheck.Checked = box.Bold;
                italicCheck.Checked = box.Italic;
                alignmentPicker.SelectedIndex = (int)box.Alignment;
                colorEditor.Text = "#" + NormalizeHex(box.ColorHex);
            }
            finally
            {
                loadingProperties = false;
            }
        }

        private void ApplyPropertyChanges()
        {
            if (loadingProperties)
                return;

            PresentationTextBox box = GetSelectedTextBox();
            if (box == null)
                return;

            box.Text = textEditor.Text ?? string.Empty;

            if (fontPicker.SelectedItem != null)
                box.FontFamily = fontPicker.SelectedItem.ToString();

            box.FontSizePoints = (float)fontSize.Value;
            box.Bold = boldCheck.Checked;
            box.Italic = italicCheck.Checked;

            if (alignmentPicker.SelectedIndex >= 0 && alignmentPicker.SelectedIndex <= 2)
                box.Alignment = (PresentationTextAlignment)alignmentPicker.SelectedIndex;

            string hex = NormalizeHex(colorEditor.Text);
            if (hex.Length == 6)
                box.ColorHex = hex;

            session.MarkDirty();
            canvas.Invalidate();
            UpdateStatus();
        }

        private void SelectFontFamily(string name)
        {
            if (string.IsNullOrEmpty(name))
                return;

            for (int i = 0; i < fontPicker.Items.Count; i++)
            {
                if (string.Equals(
                        fontPicker.Items[i].ToString(),
                        name,
                        StringComparison.CurrentCultureIgnoreCase))
                {
                    fontPicker.SelectedIndex = i;
                    return;
                }
            }
        }

        private void SaveDocument(bool forceSaveAs)
        {
            try
            {
                if (forceSaveAs || string.IsNullOrEmpty(session.FilePath))
                {
                    using (SaveFileDialog dialog = new SaveFileDialog())
                    {
                        dialog.Filter = "PowerPoint Open XML Presentation (*.pptx)|*.pptx";
                        dialog.DefaultExt = "pptx";
                        dialog.AddExtension = true;
                        dialog.FileName = SafeFileName(session.Document.Title) + ".pptx";

                        if (dialog.ShowDialog(this) != DialogResult.OK)
                            return;

                        session.SaveAs(dialog.FileName);
                    }
                }
                else
                {
                    session.Save();
                }

                UpdateStatus();
                Text = Path.GetFileName(session.FilePath) + " - Editor";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Save failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void UpdateStatus()
        {
            string fileName = string.IsNullOrEmpty(session.FilePath)
                ? "Unsaved"
                : Path.GetFileName(session.FilePath);

            statusLabel.Text =
                (session.IsDirty ? "Modified  •  " : "Saved  •  ") +
                fileName +
                "  •  " +
                session.Document.Slides.Count.ToString() +
                " slide(s)";
        }

        private void OnEditorClosing(object sender, FormClosingEventArgs e)
        {
            if (!session.IsDirty)
                return;

            DialogResult result = MessageBox.Show(
                this,
                "Save changes before closing?",
                "Unsaved presentation",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (result == DialogResult.Yes)
            {
                SaveDocument(false);

                if (session.IsDirty)
                    e.Cancel = true;
            }
        }

        private static string SafeFileName(string value)
        {
            string text = string.IsNullOrEmpty(value) ? "New Presentation" : value;
            char[] invalid = Path.GetInvalidFileNameChars();

            for (int i = 0; i < invalid.Length; i++)
                text = text.Replace(invalid[i], '_');

            return text;
        }

        private static string NormalizeHex(string value)
        {
            string candidate = (value ?? string.Empty)
                .Trim()
                .TrimStart('#')
                .ToUpperInvariant();

            if (candidate.Length != 6)
                return candidate;

            int parsed;
            if (!int.TryParse(candidate, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out parsed))
                return string.Empty;

            return candidate;
        }
    }

    internal sealed class PresentationCanvas : Panel
    {
        private Rectangle slideRectangle;

        public PresentationDocument Document { get; set; }
        public int SelectedSlideIndex { get; set; }
        public int SelectedTextBoxIndex { get; set; }
        public event EventHandler SelectionChanged;

        public PresentationCanvas()
        {
            DoubleBuffered = true;
            SelectedSlideIndex = 0;
            SelectedTextBoxIndex = -1;
            Cursor = Cursors.Default;

            MouseDown += OnCanvasMouseDown;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (Document == null ||
                SelectedSlideIndex < 0 ||
                SelectedSlideIndex >= Document.Slides.Count)
            {
                return;
            }

            PresentationSlide slide = Document.Slides[SelectedSlideIndex];
            slideRectangle = CalculateSlideRectangle();

            e.Graphics.FillRectangle(Brushes.White, slideRectangle);
            using (Pen border = new Pen(Color.FromArgb(75, 80, 88)))
                e.Graphics.DrawRectangle(border, slideRectangle);

            if (slide == null)
                return;

            for (int i = 0; i < slide.TextBoxes.Count; i++)
                DrawTextBox(e.Graphics, slide.TextBoxes[i], i);
        }

        private Rectangle CalculateSlideRectangle()
        {
            int availableWidth = Math.Max(100, ClientSize.Width - 48);
            int availableHeight = Math.Max(100, ClientSize.Height - 48);

            long widthEmu = Document != null && Document.WidthEmu > 0
                ? Document.WidthEmu
                : PresentationDocument.DefaultWidthEmu;
            long heightEmu = Document != null && Document.HeightEmu > 0
                ? Document.HeightEmu
                : PresentationDocument.DefaultHeightEmu;

            float ratio = widthEmu / (float)heightEmu;
            int width = availableWidth;
            int height = (int)Math.Round(width / ratio);

            if (height > availableHeight)
            {
                height = availableHeight;
                width = (int)Math.Round(height * ratio);
            }

            int x = (ClientSize.Width - width) / 2;
            int y = (ClientSize.Height - height) / 2;
            return new Rectangle(x, y, Math.Max(1, width), Math.Max(1, height));
        }

        private RectangleF GetTextBoxRectangle(PresentationTextBox box)
        {
            long widthEmu = Document.WidthEmu > 0
                ? Document.WidthEmu
                : PresentationDocument.DefaultWidthEmu;
            long heightEmu = Document.HeightEmu > 0
                ? Document.HeightEmu
                : PresentationDocument.DefaultHeightEmu;

            float x = slideRectangle.Left +
                (box.X / (float)widthEmu) * slideRectangle.Width;
            float y = slideRectangle.Top +
                (box.Y / (float)heightEmu) * slideRectangle.Height;
            float width = (box.Width / (float)widthEmu) * slideRectangle.Width;
            float height = (box.Height / (float)heightEmu) * slideRectangle.Height;

            return new RectangleF(x, y, Math.Max(1f, width), Math.Max(1f, height));
        }

        private void DrawTextBox(Graphics graphics, PresentationTextBox box, int index)
        {
            RectangleF rect = GetTextBoxRectangle(box);
            FontStyle style = FontStyle.Regular;

            if (box.Bold)
                style |= FontStyle.Bold;
            if (box.Italic)
                style |= FontStyle.Italic;

            float scale = slideRectangle.Width / 960f;
            float displaySize = Math.Max(6f, box.FontSizePoints * scale);

            using (Font font = SafeFont(box.FontFamily, displaySize, style))
            using (Brush brush = new SolidBrush(ParseColor(box.ColorHex)))
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

            if (index == SelectedTextBoxIndex)
            {
                using (Pen selection = new Pen(ApplicationTheme.Accent, 2f))
                    graphics.DrawRectangle(
                        selection,
                        Rectangle.Round(rect));
            }
        }

        private void OnCanvasMouseDown(object sender, MouseEventArgs e)
        {
            if (Document == null ||
                SelectedSlideIndex < 0 ||
                SelectedSlideIndex >= Document.Slides.Count)
            {
                return;
            }

            PresentationSlide slide = Document.Slides[SelectedSlideIndex];
            int found = -1;

            for (int i = slide.TextBoxes.Count - 1; i >= 0; i--)
            {
                if (GetTextBoxRectangle(slide.TextBoxes[i]).Contains(e.Location))
                {
                    found = i;
                    break;
                }
            }

            if (found != SelectedTextBoxIndex)
            {
                SelectedTextBoxIndex = found;
                Invalidate();

                if (SelectionChanged != null)
                    SelectionChanged(this, EventArgs.Empty);
            }
        }

        private static Font SafeFont(string family, float size, FontStyle style)
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
                return new Font(SystemFonts.MessageBoxFont.FontFamily, size, style);
            }
        }

        private static Color ParseColor(string value)
        {
            string candidate = (value ?? string.Empty)
                .Trim()
                .TrimStart('#');

            int parsed;
            if (candidate.Length == 6 &&
                int.TryParse(candidate, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out parsed))
            {
                return Color.FromArgb(
                    (parsed >> 16) & 0xFF,
                    (parsed >> 8) & 0xFF,
                    parsed & 0xFF);
            }

            return Color.FromArgb(32, 36, 42);
        }
    }
}
