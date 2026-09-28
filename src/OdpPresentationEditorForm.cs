using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class OdpPresentationEditSession
    {
        public PresentationDocument Document { get; private set; }
        public string FilePath { get; private set; }
        public bool IsDirty { get; private set; }

        private OdpPresentationEditSession(PresentationDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            Document = document;
            FilePath = string.Empty;
            IsDirty = true;
        }

        public static OdpPresentationEditSession CreateNew(string title)
        {
            return new OdpPresentationEditSession(
                PresentationDocument.CreateNew(title));
        }

        public static OdpPresentationEditSession Open(string path)
        {
            OdpEditSafetyResult safety = OdpEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
            {
                throw new InvalidOperationException(
                    safety == null || string.IsNullOrEmpty(safety.Warning)
                        ? "This ODP cannot yet be edited without risking unsupported-content loss."
                        : safety.Warning);
            }

            OdpPresentationEditSession session =
                new OdpPresentationEditSession(OdpReader.Read(path));
            session.FilePath = Path.GetFullPath(path);
            session.IsDirty = false;
            return session;
        }

        public void MarkDirty()
        {
            IsDirty = true;
        }

        public void Save()
        {
            if (string.IsNullOrEmpty(FilePath))
                throw new InvalidOperationException("Save As is required for a new ODP presentation.");

            ValidateWritableModel();
            OdpWriter.Save(Document, FilePath);
            IsDirty = false;
        }

        public void SaveAs(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("A destination path is required.", "path");

            if (!string.Equals(
                    Path.GetExtension(path),
                    ".odp",
                    StringComparison.OrdinalIgnoreCase))
            {
                path = Path.ChangeExtension(path, "odp");
            }

            ValidateWritableModel();
            OdpWriter.Save(Document, path);
            FilePath = Path.GetFullPath(path);
            IsDirty = false;
        }

        private void ValidateWritableModel()
        {
            for (int slideIndex = 0; slideIndex < Document.Slides.Count; slideIndex++)
            {
                PresentationSlide slide = Document.Slides[slideIndex];
                if (slide == null)
                    continue;

                for (int shapeIndex = 0; shapeIndex < slide.Shapes.Count; shapeIndex++)
                {
                    PresentationShape shape = slide.Shapes[shapeIndex];
                    if (shape == null)
                        continue;

                    if (shape.Kind != PresentationShapeKind.Rectangle &&
                        shape.Kind != PresentationShapeKind.Ellipse)
                    {
                        throw new InvalidOperationException(
                            "ODP save was stopped because slide " +
                            (slideIndex + 1).ToString() +
                            " contains a shape that the current ODP writer cannot preserve exactly: " +
                            shape.Kind.ToString() +
                            ". Only Rectangle and Ellipse are enabled until ODF custom-shape mapping is implemented.");
                    }
                }
            }
        }
    }

    internal sealed class OdpPresentationEditorForm : Form
    {
        private readonly OdpPresentationEditSession session;
        private readonly ListBox slideList;
        private readonly AdvancedPresentationCanvas canvas;
        private readonly Label statusLabel;
        private readonly Label objectTypeLabel;
        private readonly TextBox textEditor;
        private readonly ComboBox fontPicker;
        private readonly NumericUpDown fontSize;
        private readonly CheckBox boldCheck;
        private readonly CheckBox italicCheck;
        private readonly ComboBox alignmentPicker;
        private readonly TextBox textColorEditor;
        private readonly ComboBox shapeKindPicker;
        private readonly TextBox fillColorEditor;
        private readonly TextBox lineColorEditor;
        private bool loadingProperties;

        public string SavedFilePath
        {
            get { return session.FilePath; }
        }

        public OdpPresentationEditorForm(OdpPresentationEditSession editSession)
        {
            if (editSession == null)
                throw new ArgumentNullException("editSession");

            session = editSession;

            Text = "ODP Presentation Editor";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1420;
            Height = 860;
            MinimumSize = new Size(1060, 680);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);
            KeyPreview = true;

            Panel toolbar = new Panel();
            toolbar.Dock = DockStyle.Top;
            toolbar.Height = 48;
            toolbar.Padding = new Padding(8);
            toolbar.BackColor = ApplicationTheme.Toolbar;
            Controls.Add(toolbar);

            Button save = MakeButton("Save", 8, 66);
            Button saveAs = MakeButton("Save As", 80, 76);
            Button addSlide = MakeButton("+ Slide", 166, 72);
            Button deleteSlide = MakeButton("- Slide", 244, 72);
            Button addText = MakeButton("Text", 322, 64);
            Button addImage = MakeButton("Image", 392, 68);
            Button addShape = MakeButton("Shape", 466, 68);
            Button deleteObject = MakeButton("Delete object", 540, 96);

            toolbar.Controls.Add(save);
            toolbar.Controls.Add(saveAs);
            toolbar.Controls.Add(addSlide);
            toolbar.Controls.Add(deleteSlide);
            toolbar.Controls.Add(addText);
            toolbar.Controls.Add(addImage);
            toolbar.Controls.Add(addShape);
            toolbar.Controls.Add(deleteObject);

            ComboBox quickShape = new ComboBox();
            quickShape.Left = 644;
            quickShape.Top = 10;
            quickShape.Width = 130;
            quickShape.DropDownStyle = ComboBoxStyle.DropDownList;
            AddShapeKinds(quickShape);
            quickShape.SelectedIndex = 0;
            toolbar.Controls.Add(quickShape);

            Label formatBadge = new Label();
            formatBadge.Text = "ODF · ODP";
            formatBadge.Left = 790;
            formatBadge.Top = 16;
            formatBadge.AutoSize = true;
            formatBadge.ForeColor = ApplicationTheme.Accent;
            toolbar.Controls.Add(formatBadge);

            save.Click += delegate { SaveDocument(false); };
            saveAs.Click += delegate { SaveDocument(true); };
            addSlide.Click += delegate { AddSlide(); };
            deleteSlide.Click += delegate { DeleteSlide(); };
            addText.Click += delegate { AddTextBox(); };
            addImage.Click += delegate { AddImage(); };
            addShape.Click += delegate
            {
                AddShape(ShapeKindFromPickerIndex(quickShape.SelectedIndex));
            };
            deleteObject.Click += delegate { DeleteSelectedObject(); };

            statusLabel = new Label();
            statusLabel.Dock = DockStyle.Bottom;
            statusLabel.Height = 28;
            statusLabel.Padding = new Padding(10, 0, 10, 0);
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusLabel.BackColor = ApplicationTheme.Toolbar;
            statusLabel.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(statusLabel);

            Panel properties = new Panel();
            properties.Dock = DockStyle.Right;
            properties.Width = 300;
            properties.Padding = new Padding(14);
            properties.BackColor = ApplicationTheme.Sidebar;
            Controls.Add(properties);

            Label propertyTitle = MakeLabel("Object properties", 0, 0, 260, 26, true);
            properties.Controls.Add(propertyTitle);

            objectTypeLabel = MakeLabel("No object selected", 0, 28, 260, 24, false);
            objectTypeLabel.ForeColor = ApplicationTheme.SecondaryText;
            properties.Controls.Add(objectTypeLabel);

            properties.Controls.Add(MakeLabel("Text", 0, 60, 260, 20, false));
            textEditor = new TextBox();
            textEditor.Left = 0;
            textEditor.Top = 82;
            textEditor.Width = 270;
            textEditor.Height = 108;
            textEditor.Multiline = true;
            textEditor.ScrollBars = ScrollBars.Vertical;
            textEditor.BackColor = Color.FromArgb(245, 246, 248);
            textEditor.ForeColor = Color.FromArgb(28, 31, 36);
            properties.Controls.Add(textEditor);

            properties.Controls.Add(MakeLabel("Font", 0, 200, 260, 20, false));
            fontPicker = new ComboBox();
            fontPicker.Left = 0;
            fontPicker.Top = 222;
            fontPicker.Width = 270;
            fontPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            fontPicker.BackColor = Color.White;
            properties.Controls.Add(fontPicker);

            List<string> fonts = SystemFontCatalog.GetInstalledFamilyNames();
            for (int i = 0; i < fonts.Count; i++)
                fontPicker.Items.Add(fonts[i]);

            properties.Controls.Add(MakeLabel("Size", 0, 254, 70, 20, false));
            fontSize = new NumericUpDown();
            fontSize.Left = 0;
            fontSize.Top = 276;
            fontSize.Width = 82;
            fontSize.Minimum = 1;
            fontSize.Maximum = 400;
            fontSize.DecimalPlaces = 1;
            fontSize.Increment = 0.5M;
            properties.Controls.Add(fontSize);

            boldCheck = new CheckBox();
            boldCheck.Text = "Bold";
            boldCheck.Left = 96;
            boldCheck.Top = 278;
            boldCheck.Width = 68;
            boldCheck.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(boldCheck);

            italicCheck = new CheckBox();
            italicCheck.Text = "Italic";
            italicCheck.Left = 172;
            italicCheck.Top = 278;
            italicCheck.Width = 72;
            italicCheck.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(italicCheck);

            properties.Controls.Add(MakeLabel("Alignment", 0, 312, 100, 20, false));
            alignmentPicker = new ComboBox();
            alignmentPicker.Left = 0;
            alignmentPicker.Top = 334;
            alignmentPicker.Width = 126;
            alignmentPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            alignmentPicker.Items.Add("Left");
            alignmentPicker.Items.Add("Center");
            alignmentPicker.Items.Add("Right");
            properties.Controls.Add(alignmentPicker);

            properties.Controls.Add(MakeLabel("Text color", 142, 312, 100, 20, false));
            textColorEditor = new TextBox();
            textColorEditor.Left = 142;
            textColorEditor.Top = 334;
            textColorEditor.Width = 128;
            textColorEditor.MaxLength = 7;
            properties.Controls.Add(textColorEditor);

            properties.Controls.Add(MakeLabel("Shape", 0, 378, 100, 20, false));
            shapeKindPicker = new ComboBox();
            shapeKindPicker.Left = 0;
            shapeKindPicker.Top = 400;
            shapeKindPicker.Width = 270;
            shapeKindPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            AddShapeKinds(shapeKindPicker);
            properties.Controls.Add(shapeKindPicker);

            properties.Controls.Add(MakeLabel("Fill", 0, 434, 80, 20, false));
            fillColorEditor = new TextBox();
            fillColorEditor.Left = 0;
            fillColorEditor.Top = 456;
            fillColorEditor.Width = 126;
            fillColorEditor.MaxLength = 7;
            properties.Controls.Add(fillColorEditor);

            properties.Controls.Add(MakeLabel("Line", 142, 434, 80, 20, false));
            lineColorEditor = new TextBox();
            lineColorEditor.Left = 142;
            lineColorEditor.Top = 456;
            lineColorEditor.Width = 128;
            lineColorEditor.MaxLength = 7;
            properties.Controls.Add(lineColorEditor);

            Label hint = MakeLabel(
                "Drag objects to move and resize. The current ODP writer enables Rectangle and Ellipse only; unsupported shape kinds are blocked from saving instead of being silently downgraded.",
                0,
                500,
                270,
                92,
                false);
            hint.ForeColor = ApplicationTheme.SecondaryText;
            properties.Controls.Add(hint);

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
            canvasHost.Padding = new Padding(22);
            canvasHost.BackColor = ApplicationTheme.Canvas;
            Controls.Add(canvasHost);
            canvasHost.BringToFront();

            canvas = new AdvancedPresentationCanvas();
            canvas.Dock = DockStyle.Fill;
            canvas.Document = session.Document;
            canvas.BackColor = ApplicationTheme.Canvas;
            canvasHost.Controls.Add(canvas);

            slideList.SelectedIndexChanged += delegate
            {
                canvas.Document = session.Document;
                canvas.SelectedSlideIndex = slideList.SelectedIndex;
                canvas.ClearSelection();
                canvas.Invalidate();
                LoadSelectedProperties();
            };

            canvas.SelectionChanged += delegate { LoadSelectedProperties(); };
            canvas.ObjectChanged += delegate
            {
                session.MarkDirty();
                LoadSelectedProperties();
                UpdateStatus();
            };
            canvas.DeleteRequested += delegate { DeleteSelectedObject(); };

            textEditor.TextChanged += delegate { ApplyProperties(); };
            fontPicker.SelectedIndexChanged += delegate { ApplyProperties(); };
            fontSize.ValueChanged += delegate { ApplyProperties(); };
            boldCheck.CheckedChanged += delegate { ApplyProperties(); };
            italicCheck.CheckedChanged += delegate { ApplyProperties(); };
            alignmentPicker.SelectedIndexChanged += delegate { ApplyProperties(); };
            textColorEditor.TextChanged += delegate { ApplyProperties(); };
            shapeKindPicker.SelectedIndexChanged += delegate { ApplyProperties(); };
            fillColorEditor.TextChanged += delegate { ApplyProperties(); };
            lineColorEditor.TextChanged += delegate { ApplyProperties(); };

            FormClosing += OnEditorClosing;

            RefreshSlides();
            if (slideList.Items.Count > 0)
                slideList.SelectedIndex = 0;
            UpdateStatus();
        }

        public static OdpPresentationEditorForm CreateNew(string title)
        {
            return new OdpPresentationEditorForm(
                OdpPresentationEditSession.CreateNew(title));
        }

        public static OdpPresentationEditorForm Open(string path)
        {
            return new OdpPresentationEditorForm(
                OdpPresentationEditSession.Open(path));
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                SaveDocument(false);
                return true;
            }
            if (keyData == (Keys.Control | Keys.Shift | Keys.S))
            {
                SaveDocument(true);
                return true;
            }
            if (keyData == Keys.Delete)
            {
                DeleteSelectedObject();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private Button MakeButton(string text, int left, int width)
        {
            Button button = new Button();
            button.Text = text;
            button.Left = left;
            button.Top = 8;
            button.Width = width;
            button.Height = 32;
            ApplicationTheme.ApplyButton(button);
            return button;
        }

        private Label MakeLabel(string text, int left, int top, int width, int height, bool bold)
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

        private static void AddShapeKinds(ComboBox combo)
        {
            combo.Items.Clear();
            combo.Items.Add("Rectangle");
            combo.Items.Add("Ellipse");
        }

        private static PresentationShapeKind ShapeKindFromPickerIndex(int index)
        {
            return index == 1
                ? PresentationShapeKind.Ellipse
                : PresentationShapeKind.Rectangle;
        }

        private static int PickerIndexFromShapeKind(PresentationShapeKind kind)
        {
            return kind == PresentationShapeKind.Ellipse ? 1 : 0;
        }

        private void AddSlide()
        {
            PresentationSlide slide = session.Document.AddSlide("New Slide");
            slide.Name = "Slide " + session.Document.Slides.Count.ToString();
            session.MarkDirty();
            RefreshSlides();
            slideList.SelectedIndex = session.Document.Slides.Count - 1;
            UpdateStatus();
        }

        private void DeleteSlide()
        {
            int index = slideList.SelectedIndex;
            if (!session.Document.RemoveSlide(index))
            {
                MessageBox.Show(
                    this,
                    "A presentation must keep at least one slide.",
                    "Delete slide",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            session.MarkDirty();
            RefreshSlides();
            slideList.SelectedIndex = Math.Min(index, slideList.Items.Count - 1);
            UpdateStatus();
        }

        private void AddTextBox()
        {
            PresentationSlide slide = CurrentSlide();
            if (slide == null)
                return;

            PresentationTextBox box = slide.AddTextBox("New text");
            box.X = 1828800;
            box.Y = 2514600;
            box.Width = 8534400;
            box.Height = 1143000;
            box.Alignment = PresentationTextAlignment.Center;
            session.MarkDirty();
            canvas.SelectObject(EditorObjectKind.TextBox, slide.TextBoxes.Count - 1);
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private void AddShape(PresentationShapeKind kind)
        {
            PresentationSlide slide = CurrentSlide();
            if (slide == null)
                return;

            if (kind != PresentationShapeKind.Rectangle &&
                kind != PresentationShapeKind.Ellipse)
            {
                throw new NotSupportedException(
                    "The current ODP writer only enables Rectangle and Ellipse until ODF custom-shape mapping is implemented.");
            }

            PresentationShape shape = slide.AddShape(kind);
            shape.Name = kind.ToString() + " " + slide.Shapes.Count.ToString();
            session.MarkDirty();
            canvas.SelectObject(EditorObjectKind.Shape, slide.Shapes.Count - 1);
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private void AddImage()
        {
            PresentationSlide slide = CurrentSlide();
            if (slide == null)
                return;

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Images (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp";
                dialog.CheckFileExists = true;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                FileInfo info = new FileInfo(dialog.FileName);
                if (info.Length > 50L * 1024L * 1024L)
                    throw new InvalidOperationException("Image files larger than 50 MB are not accepted by the experimental editor.");

                byte[] data = File.ReadAllBytes(dialog.FileName);
                string extension = Path.GetExtension(dialog.FileName).TrimStart('.').ToLowerInvariant();
                string contentType = ImageContentType(extension);
                PresentationImage image = slide.AddImage(data, extension, contentType);
                image.Name = Path.GetFileName(dialog.FileName);

                try
                {
                    using (Image decoded = Image.FromFile(dialog.FileName))
                    {
                        long maxWidth = 7315200;
                        long maxHeight = 4114800;
                        double scale = Math.Min(
                            maxWidth / (double)Math.Max(1, decoded.Width),
                            maxHeight / (double)Math.Max(1, decoded.Height));
                        image.Width = Math.Max(914400L, (long)Math.Round(decoded.Width * scale));
                        image.Height = Math.Max(514350L, (long)Math.Round(decoded.Height * scale));
                        image.X = (PresentationDocument.DefaultWidthEmu - image.Width) / 2;
                        image.Y = (PresentationDocument.DefaultHeightEmu - image.Height) / 2;
                    }
                }
                catch { }

                session.MarkDirty();
                canvas.SelectObject(EditorObjectKind.Image, slide.Images.Count - 1);
                canvas.Invalidate();
                LoadSelectedProperties();
                UpdateStatus();
            }
        }

        private void DeleteSelectedObject()
        {
            PresentationSlide slide = CurrentSlide();
            if (slide == null || canvas.SelectedObjectIndex < 0)
                return;

            bool removed = false;
            int index = canvas.SelectedObjectIndex;

            if (canvas.SelectedObjectKind == EditorObjectKind.TextBox && index < slide.TextBoxes.Count)
            {
                slide.TextBoxes.RemoveAt(index);
                removed = true;
            }
            else if (canvas.SelectedObjectKind == EditorObjectKind.Shape && index < slide.Shapes.Count)
            {
                slide.Shapes.RemoveAt(index);
                removed = true;
            }
            else if (canvas.SelectedObjectKind == EditorObjectKind.Image && index < slide.Images.Count)
            {
                slide.Images.RemoveAt(index);
                removed = true;
            }

            if (removed)
            {
                session.MarkDirty();
                canvas.ClearSelection();
                canvas.Invalidate();
                LoadSelectedProperties();
                UpdateStatus();
            }
        }

        private PresentationSlide CurrentSlide()
        {
            int index = slideList.SelectedIndex;
            if (index < 0 || index >= session.Document.Slides.Count)
                return null;
            return session.Document.Slides[index];
        }

        private PresentationTextBox SelectedTextBox()
        {
            PresentationSlide slide = CurrentSlide();
            int index = canvas.SelectedObjectIndex;
            if (slide == null || canvas.SelectedObjectKind != EditorObjectKind.TextBox ||
                index < 0 || index >= slide.TextBoxes.Count)
                return null;
            return slide.TextBoxes[index];
        }

        private PresentationShape SelectedShape()
        {
            PresentationSlide slide = CurrentSlide();
            int index = canvas.SelectedObjectIndex;
            if (slide == null || canvas.SelectedObjectKind != EditorObjectKind.Shape ||
                index < 0 || index >= slide.Shapes.Count)
                return null;
            return slide.Shapes[index];
        }

        private void LoadSelectedProperties()
        {
            loadingProperties = true;
            try
            {
                PresentationTextBox text = SelectedTextBox();
                PresentationShape shape = SelectedShape();

                bool textEnabled = text != null;
                bool shapeEnabled = shape != null;

                objectTypeLabel.Text = textEnabled
                    ? "Text box"
                    : shapeEnabled
                        ? "Shape"
                        : canvas.SelectedObjectKind == EditorObjectKind.Image
                            ? "Image"
                            : "No object selected";

                textEditor.Enabled = textEnabled;
                fontPicker.Enabled = textEnabled;
                fontSize.Enabled = textEnabled;
                boldCheck.Enabled = textEnabled;
                italicCheck.Enabled = textEnabled;
                alignmentPicker.Enabled = textEnabled;
                textColorEditor.Enabled = textEnabled;
                shapeKindPicker.Enabled = shapeEnabled;
                fillColorEditor.Enabled = shapeEnabled;
                lineColorEditor.Enabled = shapeEnabled;

                if (textEnabled)
                {
                    textEditor.Text = text.Text ?? string.Empty;
                    SelectFont(text.FontFamily);
                    fontSize.Value = (decimal)Math.Max(1f, Math.Min(400f, text.FontSizePoints));
                    boldCheck.Checked = text.Bold;
                    italicCheck.Checked = text.Italic;
                    alignmentPicker.SelectedIndex = (int)text.Alignment;
                    textColorEditor.Text = "#" + NormalizeHex(text.ColorHex);
                }
                else
                {
                    textEditor.Text = string.Empty;
                    textColorEditor.Text = string.Empty;
                }

                if (shapeEnabled)
                {
                    shapeKindPicker.SelectedIndex = PickerIndexFromShapeKind(shape.Kind);
                    fillColorEditor.Text = "#" + NormalizeHex(shape.FillColorHex);
                    lineColorEditor.Text = "#" + NormalizeHex(shape.LineColorHex);
                }
                else
                {
                    fillColorEditor.Text = string.Empty;
                    lineColorEditor.Text = string.Empty;
                }
            }
            finally
            {
                loadingProperties = false;
            }
        }

        private void ApplyProperties()
        {
            if (loadingProperties)
                return;

            PresentationTextBox text = SelectedTextBox();
            PresentationShape shape = SelectedShape();

            if (text != null)
            {
                text.Text = textEditor.Text ?? string.Empty;
                if (fontPicker.SelectedItem != null)
                    text.FontFamily = fontPicker.SelectedItem.ToString();
                text.FontSizePoints = (float)fontSize.Value;
                text.Bold = boldCheck.Checked;
                text.Italic = italicCheck.Checked;
                if (alignmentPicker.SelectedIndex >= 0 && alignmentPicker.SelectedIndex <= 2)
                    text.Alignment = (PresentationTextAlignment)alignmentPicker.SelectedIndex;
                string color = NormalizeHex(textColorEditor.Text);
                if (color.Length == 6)
                    text.ColorHex = color;
            }

            if (shape != null)
            {
                if (shapeKindPicker.SelectedIndex >= 0 && shapeKindPicker.SelectedIndex <= 1)
                    shape.Kind = ShapeKindFromPickerIndex(shapeKindPicker.SelectedIndex);
                string fill = NormalizeHex(fillColorEditor.Text);
                string line = NormalizeHex(lineColorEditor.Text);
                if (fill.Length == 6) shape.FillColorHex = fill;
                if (line.Length == 6) shape.LineColorHex = line;
            }

            if (text != null || shape != null)
            {
                session.MarkDirty();
                canvas.Invalidate();
                UpdateStatus();
            }
        }

        private void SelectFont(string family)
        {
            for (int i = 0; i < fontPicker.Items.Count; i++)
            {
                if (string.Equals(
                        fontPicker.Items[i].ToString(),
                        family,
                        StringComparison.CurrentCultureIgnoreCase))
                {
                    fontPicker.SelectedIndex = i;
                    return;
                }
            }
        }

        private void RefreshSlides()
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

        private void SaveDocument(bool saveAs)
        {
            try
            {
                if (saveAs || string.IsNullOrEmpty(session.FilePath))
                {
                    using (SaveFileDialog dialog = new SaveFileDialog())
                    {
                        dialog.Filter = "OpenDocument Presentation (*.odp)|*.odp";
                        dialog.DefaultExt = "odp";
                        dialog.AddExtension = true;
                        dialog.FileName = SafeFileName(session.Document.Title) + ".odp";

                        if (dialog.ShowDialog(this) != DialogResult.OK)
                            return;

                        session.SaveAs(dialog.FileName);
                    }
                }
                else
                {
                    session.Save();
                }

                Text = Path.GetFileName(session.FilePath) + " - ODP Editor";
                UpdateStatus();
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
            statusLabel.Text =
                (session.IsDirty ? "Modified  •  " : "Saved  •  ") +
                (string.IsNullOrEmpty(session.FilePath)
                    ? "Unsaved ODP"
                    : Path.GetFileName(session.FilePath)) +
                "  •  " + session.Document.Slides.Count.ToString() + " slide(s)";
        }

        private void OnEditorClosing(object sender, FormClosingEventArgs e)
        {
            if (!session.IsDirty)
                return;

            DialogResult result = MessageBox.Show(
                this,
                "Save changes before closing?",
                "Unsaved ODP presentation",
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

        private static string ImageContentType(string extension)
        {
            string value = (extension ?? string.Empty).ToLowerInvariant();
            if (value == "png") return "image/png";
            if (value == "jpg" || value == "jpeg") return "image/jpeg";
            if (value == "gif") return "image/gif";
            if (value == "bmp") return "image/bmp";
            throw new NotSupportedException("Unsupported image type: " + value);
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
            return int.TryParse(
                    candidate,
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture,
                    out parsed)
                ? candidate
                : string.Empty;
        }

        private static string SafeFileName(string value)
        {
            string text = string.IsNullOrEmpty(value) ? "New Presentation" : value;
            char[] invalid = Path.GetInvalidFileNameChars();
            for (int i = 0; i < invalid.Length; i++)
                text = text.Replace(invalid[i], '_');
            return text;
        }
    }
}
