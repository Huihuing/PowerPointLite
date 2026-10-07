using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal enum EditorObjectKind
    {
        None,
        TextBox,
        Shape,
        Image
    }

    internal sealed class AdvancedPresentationEditorForm : Form
    {
        private readonly PresentationEditSession session;
        private readonly PresentationHistory history;
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
        private readonly Button undoButton;
        private readonly Button redoButton;
        private readonly Button deleteObjectButton;
        private readonly Button copyButton;
        private readonly Button pasteButton;

        private bool loadingProperties;
        private bool propertyEditSnapshotActive;
        private PresentationTextBox copiedTextBox;
        private PresentationShape copiedShape;
        private PresentationImage copiedImage;

        public string SavedFilePath
        {
            get { return session.FilePath; }
        }

        public AdvancedPresentationEditorForm(PresentationEditSession editSession)
        {
            if (editSession == null)
                throw new ArgumentNullException("editSession");

            session = editSession;
            history = new PresentationHistory(40);

            Text = "Presentation Editor";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1420;
            Height = 860;
            MinimumSize = new Size(1060, 680);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);
            KeyPreview = true;

            Panel mainToolbar = new Panel();
            mainToolbar.Dock = DockStyle.Top;
            mainToolbar.Height = 48;
            mainToolbar.Padding = new Padding(8, 8, 8, 8);
            mainToolbar.BackColor = ApplicationTheme.Toolbar;
            Controls.Add(mainToolbar);

            Button saveButton = MakeButton("Save", 8, 66);
            Button saveAsButton = MakeButton("Save As", 80, 76);
            undoButton = MakeButton("Undo", 166, 64);
            redoButton = MakeButton("Redo", 236, 64);
            Button addTextButton = MakeButton("Text", 316, 64);
            Button addImageButton = MakeButton("Image", 386, 68);
            Button addShapeButton = MakeButton("Shape", 460, 68);
            deleteObjectButton = MakeButton("Delete", 534, 70);
            copyButton = MakeButton("Copy", 614, 64);
            pasteButton = MakeButton("Paste", 684, 64);

            mainToolbar.Controls.Add(saveButton);
            mainToolbar.Controls.Add(saveAsButton);
            mainToolbar.Controls.Add(undoButton);
            mainToolbar.Controls.Add(redoButton);
            mainToolbar.Controls.Add(addTextButton);
            mainToolbar.Controls.Add(addImageButton);
            mainToolbar.Controls.Add(addShapeButton);
            mainToolbar.Controls.Add(deleteObjectButton);
            mainToolbar.Controls.Add(copyButton);
            mainToolbar.Controls.Add(pasteButton);

            ComboBox quickShapeKind = new ComboBox();
            quickShapeKind.Top = 10;
            quickShapeKind.Width = 148;
            quickShapeKind.DropDownStyle = ComboBoxStyle.DropDownList;
            quickShapeKind.FlatStyle = FlatStyle.Flat;
            quickShapeKind.BackColor = ApplicationTheme.Surface;
            quickShapeKind.ForeColor = ApplicationTheme.PrimaryText;
            AddShapeKindItems(quickShapeKind);
            quickShapeKind.SelectedIndex = 0;
            mainToolbar.Controls.Add(quickShapeKind);

            Action layoutMainToolbar =
                delegate
                {
                    int gap = 12;
                    int minimumSelectorWidth = 104;
                    int desiredSelectorWidth = 148;
                    int selectorRight =
                        Math.Max(
                            pasteButton.Right +
                            gap,
                            mainToolbar.ClientSize.Width -
                            10);

                    int availableSelectorWidth =
                        selectorRight -
                        pasteButton.Right -
                        gap;

                    if (availableSelectorWidth <
                        minimumSelectorWidth)
                    {
                        quickShapeKind.Visible = false;
                        return;
                    }

                    quickShapeKind.Visible = true;
                    quickShapeKind.Width =
                        Math.Min(
                            desiredSelectorWidth,
                            availableSelectorWidth);
                    quickShapeKind.Left =
                        selectorRight -
                        quickShapeKind.Width;
                };

            mainToolbar.Resize += delegate
            {
                layoutMainToolbar();
            };

            layoutMainToolbar();

            saveButton.Click += delegate { SaveDocument(false); };
            saveAsButton.Click += delegate { SaveDocument(true); };
            undoButton.Click += delegate { Undo(); };
            redoButton.Click += delegate { Redo(); };
            addTextButton.Click += delegate { AddTextBox(); };
            addImageButton.Click += delegate { AddImage(); };
            addShapeButton.Click += delegate
            {
                AddShape((PresentationShapeKind)Math.Max(0, quickShapeKind.SelectedIndex));
            };
            deleteObjectButton.Click += delegate { DeleteSelectedObject(); };
            copyButton.Click += delegate { CopySelectedObject(); };
            pasteButton.Click += delegate { PasteObject(); };

            Panel slideToolbar = new Panel();
            slideToolbar.Dock = DockStyle.Top;
            slideToolbar.Height = 40;
            slideToolbar.Padding = new Padding(8, 5, 8, 5);
            slideToolbar.BackColor = ApplicationTheme.Toolbar;
            slideToolbar.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (Pen divider = new Pen(ApplicationTheme.Divider))
                {
                    e.Graphics.DrawLine(
                        divider,
                        0,
                        slideToolbar.ClientSize.Height - 1,
                        Math.Max(
                            0,
                            slideToolbar.ClientSize.Width - 1),
                        slideToolbar.ClientSize.Height - 1);
                }
            };
            Controls.Add(slideToolbar);
            slideToolbar.BringToFront();

            Button addSlide = MakeButton("+ Slide", 8, 72);
            Button deleteSlide = MakeButton("- Slide", 86, 72);
            Button moveUp = MakeButton("Up", 164, 54);
            Button moveDown = MakeButton("Down", 224, 64);
            slideToolbar.Controls.Add(addSlide);
            slideToolbar.Controls.Add(deleteSlide);
            slideToolbar.Controls.Add(moveUp);
            slideToolbar.Controls.Add(moveDown);

            addSlide.Top = 4;
            deleteSlide.Top = 4;
            moveUp.Top = 4;
            moveDown.Top = 4;

            addSlide.Click += delegate { AddSlide(); };
            deleteSlide.Click += delegate { DeleteSlide(); };
            moveUp.Click += delegate { MoveSlide(-1); };
            moveDown.Click += delegate { MoveSlide(1); };

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
            properties.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (Pen divider = new Pen(ApplicationTheme.Divider))
                {
                    e.Graphics.DrawLine(
                        divider,
                        0,
                        0,
                        0,
                        Math.Max(
                            0,
                            properties.ClientSize.Height - 1));
                }
            };
            Controls.Add(properties);

            Label title = MakeLabel("Object properties", 0, 0, 260, 26, true);
            properties.Controls.Add(title);

            objectTypeLabel = MakeLabel("No object selected", 0, 28, 260, 24, false);
            objectTypeLabel.ForeColor = ApplicationTheme.SecondaryText;
            properties.Controls.Add(objectTypeLabel);

            Label textLabel = MakeLabel("Text", 0, 60, 260, 20, false);
            properties.Controls.Add(textLabel);

            textEditor = new TextBox();
            textEditor.Left = 0;
            textEditor.Top = 82;
            textEditor.Width = 270;
            textEditor.Height = 108;
            textEditor.Multiline = true;
            textEditor.ScrollBars = ScrollBars.Vertical;
            textEditor.BackColor = ApplicationTheme.Surface;
            textEditor.ForeColor = ApplicationTheme.PrimaryText;
            textEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(textEditor);

            Label fontLabel = MakeLabel("Font", 0, 200, 260, 20, false);
            properties.Controls.Add(fontLabel);

            fontPicker = new ComboBox();
            fontPicker.Left = 0;
            fontPicker.Top = 222;
            fontPicker.Width = 270;
            fontPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            fontPicker.FlatStyle = FlatStyle.Flat;
            fontPicker.BackColor = ApplicationTheme.Surface;
            fontPicker.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(fontPicker);

            List<string> fonts = SystemFontCatalog.GetInstalledFamilyNames();
            for (int i = 0; i < fonts.Count; i++)
                fontPicker.Items.Add(fonts[i]);

            Label sizeLabel = MakeLabel("Size", 0, 254, 70, 20, false);
            properties.Controls.Add(sizeLabel);

            fontSize = new NumericUpDown();
            fontSize.Left = 0;
            fontSize.Top = 276;
            fontSize.Width = 82;
            fontSize.Minimum = 1;
            fontSize.Maximum = 400;
            fontSize.DecimalPlaces = 1;
            fontSize.Increment = 0.5M;
            fontSize.BackColor = ApplicationTheme.Surface;
            fontSize.ForeColor = ApplicationTheme.PrimaryText;
            fontSize.BorderStyle = BorderStyle.FixedSingle;
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

            Label alignLabel = MakeLabel("Alignment", 0, 312, 100, 20, false);
            properties.Controls.Add(alignLabel);

            alignmentPicker = new ComboBox();
            alignmentPicker.Left = 0;
            alignmentPicker.Top = 334;
            alignmentPicker.Width = 126;
            alignmentPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            alignmentPicker.FlatStyle = FlatStyle.Flat;
            alignmentPicker.BackColor = ApplicationTheme.Surface;
            alignmentPicker.ForeColor = ApplicationTheme.PrimaryText;
            alignmentPicker.Items.Add("Left");
            alignmentPicker.Items.Add("Center");
            alignmentPicker.Items.Add("Right");
            properties.Controls.Add(alignmentPicker);

            Label textColorLabel = MakeLabel("Text color", 142, 312, 100, 20, false);
            properties.Controls.Add(textColorLabel);

            textColorEditor = new TextBox();
            textColorEditor.Left = 142;
            textColorEditor.Top = 334;
            textColorEditor.Width = 128;
            textColorEditor.MaxLength = 7;
            textColorEditor.BackColor = ApplicationTheme.Surface;
            textColorEditor.ForeColor = ApplicationTheme.PrimaryText;
            textColorEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(textColorEditor);

            Label shapeLabel = MakeLabel("Shape", 0, 378, 100, 20, false);
            properties.Controls.Add(shapeLabel);

            shapeKindPicker = new ComboBox();
            shapeKindPicker.Left = 0;
            shapeKindPicker.Top = 400;
            shapeKindPicker.Width = 270;
            shapeKindPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            shapeKindPicker.FlatStyle = FlatStyle.Flat;
            shapeKindPicker.BackColor = ApplicationTheme.Surface;
            shapeKindPicker.ForeColor = ApplicationTheme.PrimaryText;
            AddShapeKindItems(shapeKindPicker);
            properties.Controls.Add(shapeKindPicker);

            Label fillLabel = MakeLabel("Fill", 0, 434, 80, 20, false);
            properties.Controls.Add(fillLabel);
            fillColorEditor = new TextBox();
            fillColorEditor.Left = 0;
            fillColorEditor.Top = 456;
            fillColorEditor.Width = 126;
            fillColorEditor.MaxLength = 7;
            fillColorEditor.BackColor = ApplicationTheme.Surface;
            fillColorEditor.ForeColor = ApplicationTheme.PrimaryText;
            fillColorEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(fillColorEditor);

            Label lineLabel = MakeLabel("Line", 142, 434, 80, 20, false);
            properties.Controls.Add(lineLabel);
            lineColorEditor = new TextBox();
            lineColorEditor.Left = 142;
            lineColorEditor.Top = 456;
            lineColorEditor.Width = 128;
            lineColorEditor.MaxLength = 7;
            lineColorEditor.BackColor = ApplicationTheme.Surface;
            lineColorEditor.ForeColor = ApplicationTheme.PrimaryText;
            lineColorEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(lineColorEditor);

            Label hint = MakeLabel(
                "Drag objects to move. Drag the lower-right handle to resize. Arrow keys nudge the selected object.",
                0,
                500,
                270,
                80,
                false);
            hint.ForeColor = ApplicationTheme.SecondaryText;
            properties.Controls.Add(hint);

            Panel slidePanel = new Panel();
            slidePanel.Dock = DockStyle.Left;
            slidePanel.Width = 220;
            slidePanel.Padding = new Padding(10);
            slidePanel.BackColor = ApplicationTheme.Sidebar;
            slidePanel.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (Pen divider = new Pen(ApplicationTheme.Divider))
                {
                    e.Graphics.DrawLine(
                        divider,
                        slidePanel.ClientSize.Width - 1,
                        0,
                        slidePanel.ClientSize.Width - 1,
                        Math.Max(
                            0,
                            slidePanel.ClientSize.Height - 1));
                }
            };
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
            ApplicationTheme.ApplySlideList(
                slideList);
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

            canvas.SelectionChanged += delegate
            {
                propertyEditSnapshotActive = false;
                LoadSelectedProperties();
                UpdateButtons();
            };
            canvas.TransformStarting += delegate { CaptureHistory(); };
            canvas.ObjectChanged += delegate
            {
                session.MarkDirty();
                propertyEditSnapshotActive = false;
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

            textEditor.Leave += delegate { propertyEditSnapshotActive = false; };
            fontPicker.Leave += delegate { propertyEditSnapshotActive = false; };
            fontSize.Leave += delegate { propertyEditSnapshotActive = false; };
            textColorEditor.Leave += delegate { propertyEditSnapshotActive = false; };
            shapeKindPicker.Leave += delegate { propertyEditSnapshotActive = false; };
            fillColorEditor.Leave += delegate { propertyEditSnapshotActive = false; };
            lineColorEditor.Leave += delegate { propertyEditSnapshotActive = false; };

            FormClosing += OnEditorClosing;

            RefreshSlideList();
            if (slideList.Items.Count > 0)
                slideList.SelectedIndex = 0;
            UpdateStatus();
            UpdateButtons();
        }

        public static AdvancedPresentationEditorForm CreateNew(string title)
        {
            return new AdvancedPresentationEditorForm(
                PresentationEditSession.CreateNew(title));
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

            if (keyData == (Keys.Control | Keys.Z))
            {
                Undo();
                return true;
            }

            if (keyData == (Keys.Control | Keys.Y))
            {
                Redo();
                return true;
            }

            if (keyData == (Keys.Control | Keys.C))
            {
                CopySelectedObject();
                return true;
            }

            if (keyData == (Keys.Control | Keys.V))
            {
                PasteObject();
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

        private static void AddShapeKindItems(ComboBox combo)
        {
            combo.Items.Add("Rectangle");
            combo.Items.Add("Rounded rectangle");
            combo.Items.Add("Ellipse");
            combo.Items.Add("Triangle");
            combo.Items.Add("Diamond");
        }

        private void AddSlide()
        {
            CaptureHistory();
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

            CaptureHistory();

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

        private void MoveSlide(int direction)
        {
            int from = slideList.SelectedIndex;
            int to = from + direction;
            if (from < 0 || to < 0 || to >= session.Document.Slides.Count)
                return;

            CaptureHistory();
            if (!session.MoveSlide(from, to))
                return;

            RefreshSlideList();
            slideList.SelectedIndex = to;
            UpdateStatus();
        }

        private void AddTextBox()
        {
            int slideIndex = slideList.SelectedIndex;
            if (slideIndex < 0)
                return;

            CaptureHistory();
            PresentationTextBox box = session.AddTextBox(slideIndex, "Text");
            if (box == null)
                return;

            box.Name = "Text Box " + session.Document.Slides[slideIndex].TextBoxes.Count.ToString();
            box.X = 1828800;
            box.Y = 2514600;
            box.Width = 8534400;
            box.Height = 1143000;
            box.Alignment = PresentationTextAlignment.Center;
            canvas.Document = session.Document;
            canvas.SelectObject(EditorObjectKind.TextBox,
                session.Document.Slides[slideIndex].TextBoxes.Count - 1);
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private void AddShape(PresentationShapeKind kind)
        {
            int slideIndex = slideList.SelectedIndex;
            if (slideIndex < 0)
                return;

            CaptureHistory();
            PresentationShape shape = session.AddShape(slideIndex, kind);
            if (shape == null)
                return;

            int count = session.Document.Slides[slideIndex].Shapes.Count;
            shape.Name = kind.ToString() + " " + count.ToString();
            shape.X += (count % 4) * 228600;
            shape.Y += (count % 4) * 228600;
            canvas.Document = session.Document;
            canvas.SelectObject(EditorObjectKind.Shape, count - 1);
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private void AddImage()
        {
            int slideIndex = slideList.SelectedIndex;
            if (slideIndex < 0)
                return;

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter =
                    "Supported images|*.png;*.jpg;*.jpeg;*.gif;*.bmp|PNG|*.png|JPEG|*.jpg;*.jpeg|GIF|*.gif|BMP|*.bmp";
                dialog.Multiselect = false;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                byte[] data = File.ReadAllBytes(dialog.FileName);
                string extension = Path.GetExtension(dialog.FileName)
                    .TrimStart('.')
                    .ToLowerInvariant();
                string contentType = ImageContentType(extension);

                if (string.IsNullOrEmpty(contentType))
                    return;

                CaptureHistory();
                PresentationImage item = session.AddImage(
                    slideIndex,
                    data,
                    extension,
                    contentType);
                if (item == null)
                    return;

                item.Name = Path.GetFileName(dialog.FileName);
                SetImageDefaultSize(item, dialog.FileName);
                int index = session.Document.Slides[slideIndex].Images.Count - 1;
                canvas.Document = session.Document;
                canvas.SelectObject(EditorObjectKind.Image, index);
                canvas.Invalidate();
                LoadSelectedProperties();
                UpdateStatus();
            }
        }

        private void SetImageDefaultSize(PresentationImage item, string path)
        {
            try
            {
                using (Image image = Image.FromFile(path))
                {
                    double ratio = image.Width / (double)Math.Max(1, image.Height);
                    long maxWidth = 5486400;
                    long maxHeight = 3657600;
                    long width = maxWidth;
                    long height = (long)Math.Round(width / ratio);

                    if (height > maxHeight)
                    {
                        height = maxHeight;
                        width = (long)Math.Round(height * ratio);
                    }

                    item.Width = Math.Max(914400L, width);
                    item.Height = Math.Max(914400L, height);
                    item.X = Math.Max(0L, (session.Document.WidthEmu - item.Width) / 2L);
                    item.Y = Math.Max(0L, (session.Document.HeightEmu - item.Height) / 2L);
                }
            }
            catch
            {
                item.X = 1828800;
                item.Y = 1600200;
                item.Width = 5486400;
                item.Height = 3657600;
            }
        }

        private void DeleteSelectedObject()
        {
            int slideIndex = canvas.SelectedSlideIndex;
            int index = canvas.SelectedObjectIndex;
            if (slideIndex < 0 || index < 0)
                return;

            CaptureHistory();
            bool removed = false;

            if (canvas.SelectedObjectKind == EditorObjectKind.TextBox)
                removed = session.RemoveTextBox(slideIndex, index);
            else if (canvas.SelectedObjectKind == EditorObjectKind.Shape)
                removed = session.RemoveShape(slideIndex, index);
            else if (canvas.SelectedObjectKind == EditorObjectKind.Image)
                removed = session.RemoveImage(slideIndex, index);

            if (!removed)
                return;

            canvas.Document = session.Document;
            canvas.ClearSelection();
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private void CopySelectedObject()
        {
            copiedTextBox = null;
            copiedShape = null;
            copiedImage = null;

            PresentationSlide slide = GetSelectedSlide();
            if (slide == null || canvas.SelectedObjectIndex < 0)
                return;

            int index = canvas.SelectedObjectIndex;
            if (canvas.SelectedObjectKind == EditorObjectKind.TextBox && index < slide.TextBoxes.Count)
                copiedTextBox = slide.TextBoxes[index].Clone();
            else if (canvas.SelectedObjectKind == EditorObjectKind.Shape && index < slide.Shapes.Count)
                copiedShape = slide.Shapes[index].Clone();
            else if (canvas.SelectedObjectKind == EditorObjectKind.Image && index < slide.Images.Count)
                copiedImage = slide.Images[index].Clone();

            UpdateButtons();
        }

        private void PasteObject()
        {
            PresentationSlide slide = GetSelectedSlide();
            if (slide == null)
                return;

            if (copiedTextBox == null && copiedShape == null && copiedImage == null)
                return;

            CaptureHistory();
            const long offset = 228600;

            if (copiedTextBox != null)
            {
                PresentationTextBox copy = copiedTextBox.Clone();
                copy.X += offset;
                copy.Y += offset;
                slide.TextBoxes.Add(copy);
                canvas.SelectObject(EditorObjectKind.TextBox, slide.TextBoxes.Count - 1);
            }
            else if (copiedShape != null)
            {
                PresentationShape copy = copiedShape.Clone();
                copy.X += offset;
                copy.Y += offset;
                slide.Shapes.Add(copy);
                canvas.SelectObject(EditorObjectKind.Shape, slide.Shapes.Count - 1);
            }
            else if (copiedImage != null)
            {
                PresentationImage copy = copiedImage.Clone();
                copy.X += offset;
                copy.Y += offset;
                slide.Images.Add(copy);
                canvas.SelectObject(EditorObjectKind.Image, slide.Images.Count - 1);
            }

            session.MarkDirty();
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private PresentationSlide GetSelectedSlide()
        {
            int index = canvas.SelectedSlideIndex;
            if (index < 0 || index >= session.Document.Slides.Count)
                return null;
            return session.Document.Slides[index];
        }

        private void CaptureHistory()
        {
            history.Capture(session.Document);
            UpdateButtons();
        }

        private void Undo()
        {
            PresentationDocument restored = history.Undo(session.Document);
            if (restored == null)
                return;

            session.ReplaceDocument(restored, true);
            RefreshAfterHistory();
        }

        private void Redo()
        {
            PresentationDocument restored = history.Redo(session.Document);
            if (restored == null)
                return;

            session.ReplaceDocument(restored, true);
            RefreshAfterHistory();
        }

        private void RefreshAfterHistory()
        {
            int slide = Math.Max(0, Math.Min(
                slideList.SelectedIndex,
                session.Document.Slides.Count - 1));
            canvas.Document = session.Document;
            RefreshSlideList();
            slideList.SelectedIndex = slide;
            canvas.SelectedSlideIndex = slide;
            canvas.ClearSelection();
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
            UpdateButtons();
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

        private void LoadSelectedProperties()
        {
            loadingProperties = true;

            try
            {
                PresentationTextBox text = canvas.GetSelectedTextBox();
                PresentationShape shape = canvas.GetSelectedShape();
                PresentationImage image = canvas.GetSelectedImage();

                bool textEnabled = text != null;
                bool shapeEnabled = shape != null;

                objectTypeLabel.Text = text != null
                    ? "Text box"
                    : shape != null
                        ? "Shape"
                        : image != null
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

                if (text != null)
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

                if (shape != null)
                {
                    shapeKindPicker.SelectedIndex = (int)shape.Kind;
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

            PresentationTextBox text = canvas.GetSelectedTextBox();
            PresentationShape shape = canvas.GetSelectedShape();
            if (text == null && shape == null)
                return;

            if (!propertyEditSnapshotActive)
            {
                history.Capture(session.Document);
                propertyEditSnapshotActive = true;
            }

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

                string textColor = NormalizeHex(textColorEditor.Text);
                if (textColor.Length == 6)
                    text.ColorHex = textColor;
            }

            if (shape != null)
            {
                if (shapeKindPicker.SelectedIndex >= 0 && shapeKindPicker.SelectedIndex <= 4)
                    shape.Kind = (PresentationShapeKind)shapeKindPicker.SelectedIndex;

                string fill = NormalizeHex(fillColorEditor.Text);
                string line = NormalizeHex(lineColorEditor.Text);
                if (fill.Length == 6)
                    shape.FillColorHex = fill;
                if (line.Length == 6)
                    shape.LineColorHex = line;
            }

            session.MarkDirty();
            canvas.Invalidate();
            UpdateStatus();
            UpdateButtons();
        }

        private void SelectFont(string name)
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

                Text = Path.GetFileName(session.FilePath) + " - Editor";
                propertyEditSnapshotActive = false;
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
            string fileName = string.IsNullOrEmpty(session.FilePath)
                ? "Unsaved"
                : Path.GetFileName(session.FilePath);

            statusLabel.Text =
                (session.IsDirty ? "Modified  •  " : "Saved  •  ") +
                fileName + "  •  " +
                session.Document.Slides.Count.ToString() +
                " slide(s)";
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            undoButton.Enabled = history.CanUndo;
            redoButton.Enabled = history.CanRedo;
            bool hasSelection = canvas != null &&
                canvas.SelectedObjectKind != EditorObjectKind.None &&
                canvas.SelectedObjectIndex >= 0;
            deleteObjectButton.Enabled = hasSelection;
            copyButton.Enabled = hasSelection;
            pasteButton.Enabled = copiedTextBox != null || copiedShape != null || copiedImage != null;
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

        private static string ImageContentType(string extension)
        {
            if (extension == "png")
                return "image/png";
            if (extension == "jpg" || extension == "jpeg")
                return "image/jpeg";
            if (extension == "gif")
                return "image/gif";
            if (extension == "bmp")
                return "image/bmp";
            return string.Empty;
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

    internal sealed class AdvancedPresentationCanvas : Panel
    {
        private Rectangle slideRectangle;
        private bool dragging;
        private bool resizing;
        private bool historyStarted;
        private Point dragStart;
        private long originalX;
        private long originalY;
        private long originalWidth;
        private long originalHeight;

        public PresentationDocument Document { get; set; }
        public int SelectedSlideIndex { get; set; }
        public EditorObjectKind SelectedObjectKind { get; private set; }
        public int SelectedObjectIndex { get; private set; }

        public event EventHandler SelectionChanged;
        public event EventHandler TransformStarting;
        public event EventHandler ObjectChanged;
        public event EventHandler DeleteRequested;

        public AdvancedPresentationCanvas()
        {
            DoubleBuffered = true;
            TabStop = true;
            SelectedSlideIndex = 0;
            SelectedObjectKind = EditorObjectKind.None;
            SelectedObjectIndex = -1;
            Cursor = Cursors.Default;
        }

        public void ClearSelection()
        {
            SelectedObjectKind = EditorObjectKind.None;
            SelectedObjectIndex = -1;
            Invalidate();
        }

        public void SelectObject(EditorObjectKind kind, int index)
        {
            SelectedObjectKind = kind;
            SelectedObjectIndex = index;
            Invalidate();

            if (SelectionChanged != null)
                SelectionChanged(this, EventArgs.Empty);
        }

        public PresentationTextBox GetSelectedTextBox()
        {
            PresentationSlide slide = GetSlide();
            if (slide == null ||
                SelectedObjectKind != EditorObjectKind.TextBox ||
                SelectedObjectIndex < 0 ||
                SelectedObjectIndex >= slide.TextBoxes.Count)
            {
                return null;
            }

            return slide.TextBoxes[SelectedObjectIndex];
        }

        public PresentationShape GetSelectedShape()
        {
            PresentationSlide slide = GetSlide();
            if (slide == null ||
                SelectedObjectKind != EditorObjectKind.Shape ||
                SelectedObjectIndex < 0 ||
                SelectedObjectIndex >= slide.Shapes.Count)
            {
                return null;
            }

            return slide.Shapes[SelectedObjectIndex];
        }

        public PresentationImage GetSelectedImage()
        {
            PresentationSlide slide = GetSlide();
            if (slide == null ||
                SelectedObjectKind != EditorObjectKind.Image ||
                SelectedObjectIndex < 0 ||
                SelectedObjectIndex >= slide.Images.Count)
            {
                return null;
            }

            return slide.Images[SelectedObjectIndex];
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

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            ApplicationTheme.DrawSlideSurface(
                e.Graphics,
                slideRectangle);

            if (slide == null)
                return;

            for (int i = 0; i < slide.Shapes.Count; i++)
                DrawShape(e.Graphics, slide.Shapes[i], i);

            for (int i = 0; i < slide.Images.Count; i++)
                DrawImage(e.Graphics, slide.Images[i], i);

            for (int i = 0; i < slide.TextBoxes.Count; i++)
                DrawTextBox(e.Graphics, slide.TextBoxes[i], i);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();

            if (e.Button != MouseButtons.Left)
                return;

            EditorObjectKind kind;
            int index;
            RectangleF rect;

            if (!HitTest(e.Location, out kind, out index, out rect))
            {
                ClearSelection();
                if (SelectionChanged != null)
                    SelectionChanged(this, EventArgs.Empty);
                return;
            }

            SelectObject(kind, index);
            dragging = true;
            resizing = ResizeHandle(rect).Contains(e.Location);
            historyStarted = false;
            dragStart = e.Location;
            GetSelectedGeometry(out originalX, out originalY, out originalWidth, out originalHeight);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (!dragging)
            {
                RectangleF selected = GetSelectedRectangle();
                Cursor = selected != RectangleF.Empty && ResizeHandle(selected).Contains(e.Location)
                    ? Cursors.SizeNWSE
                    : Cursors.Default;
                return;
            }

            int dx = e.X - dragStart.X;
            int dy = e.Y - dragStart.Y;
            if (dx == 0 && dy == 0)
                return;

            if (!historyStarted)
            {
                historyStarted = true;
                if (TransformStarting != null)
                    TransformStarting(this, EventArgs.Empty);
            }

            long dxEmu = PixelsToEmuX(dx);
            long dyEmu = PixelsToEmuY(dy);

            if (resizing)
            {
                SetSelectedGeometry(
                    originalX,
                    originalY,
                    Math.Max(91440L, originalWidth + dxEmu),
                    Math.Max(91440L, originalHeight + dyEmu));
            }
            else
            {
                long width = originalWidth;
                long height = originalHeight;
                long x = Math.Max(0L, Math.Min(
                    Math.Max(0L, Document.WidthEmu - width),
                    originalX + dxEmu));
                long y = Math.Max(0L, Math.Min(
                    Math.Max(0L, Document.HeightEmu - height),
                    originalY + dyEmu));
                SetSelectedGeometry(x, y, width, height);
            }

            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);

            if (e.Button != MouseButtons.Left)
                return;

            bool changed = dragging && historyStarted;
            dragging = false;
            resizing = false;
            historyStarted = false;

            if (changed && ObjectChanged != null)
                ObjectChanged(this, EventArgs.Empty);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys code = keyData & Keys.KeyCode;
            bool shift = (keyData & Keys.Shift) == Keys.Shift;

            if (code == Keys.Delete && SelectedObjectKind != EditorObjectKind.None)
            {
                if (DeleteRequested != null)
                    DeleteRequested(this, EventArgs.Empty);
                return true;
            }

            if (code == Keys.Left || code == Keys.Right ||
                code == Keys.Up || code == Keys.Down)
            {
                if (SelectedObjectKind == EditorObjectKind.None)
                    return base.ProcessCmdKey(ref msg, keyData);

                if (TransformStarting != null)
                    TransformStarting(this, EventArgs.Empty);

                long step = shift ? 228600L : 45720L;
                long x;
                long y;
                long width;
                long height;
                GetSelectedGeometry(out x, out y, out width, out height);

                if (code == Keys.Left)
                    x -= step;
                else if (code == Keys.Right)
                    x += step;
                else if (code == Keys.Up)
                    y -= step;
                else if (code == Keys.Down)
                    y += step;

                x = Math.Max(0L, Math.Min(Math.Max(0L, Document.WidthEmu - width), x));
                y = Math.Max(0L, Math.Min(Math.Max(0L, Document.HeightEmu - height), y));
                SetSelectedGeometry(x, y, width, height);
                Invalidate();

                if (ObjectChanged != null)
                    ObjectChanged(this, EventArgs.Empty);
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private PresentationSlide GetSlide()
        {
            if (Document == null ||
                SelectedSlideIndex < 0 ||
                SelectedSlideIndex >= Document.Slides.Count)
            {
                return null;
            }

            return Document.Slides[SelectedSlideIndex];
        }

        private Rectangle CalculateSlideRectangle()
        {
            int availableWidth = Math.Max(100, ClientSize.Width - 48);
            int availableHeight = Math.Max(100, ClientSize.Height - 48);
            long widthEmu = Document.WidthEmu > 0
                ? Document.WidthEmu
                : PresentationDocument.DefaultWidthEmu;
            long heightEmu = Document.HeightEmu > 0
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

            return new Rectangle(
                (ClientSize.Width - width) / 2,
                (ClientSize.Height - height) / 2,
                Math.Max(1, width),
                Math.Max(1, height));
        }

        private RectangleF ToRectangle(long x, long y, long width, long height)
        {
            float left = slideRectangle.Left +
                (x / (float)Document.WidthEmu) * slideRectangle.Width;
            float top = slideRectangle.Top +
                (y / (float)Document.HeightEmu) * slideRectangle.Height;
            float w = (width / (float)Document.WidthEmu) * slideRectangle.Width;
            float h = (height / (float)Document.HeightEmu) * slideRectangle.Height;
            return new RectangleF(left, top, Math.Max(1f, w), Math.Max(1f, h));
        }

        private void DrawTextBox(Graphics graphics, PresentationTextBox box, int index)
        {
            RectangleF rect = ToRectangle(box.X, box.Y, box.Width, box.Height);
            FontStyle style = FontStyle.Regular;
            if (box.Bold)
                style |= FontStyle.Bold;
            if (box.Italic)
                style |= FontStyle.Italic;

            float scale = slideRectangle.Width / 960f;
            float displaySize = Math.Max(6f, box.FontSizePoints * scale);

            using (Font font = SafeFont(box.FontFamily, displaySize, style))
            using (Brush brush = new SolidBrush(ParseColor(box.ColorHex, Color.FromArgb(32, 36, 42))))
            using (StringFormat format = new StringFormat())
            {
                format.Alignment = box.Alignment == PresentationTextAlignment.Center
                    ? StringAlignment.Center
                    : box.Alignment == PresentationTextAlignment.Right
                        ? StringAlignment.Far
                        : StringAlignment.Near;
                format.LineAlignment = StringAlignment.Center;
                format.Trimming = StringTrimming.EllipsisCharacter;
                graphics.DrawString(box.Text ?? string.Empty, font, brush, rect, format);
            }

            DrawSelectionIfNeeded(graphics, EditorObjectKind.TextBox, index, rect);
        }

        private void DrawShape(Graphics graphics, PresentationShape shape, int index)
        {
            RectangleF rect = ToRectangle(shape.X, shape.Y, shape.Width, shape.Height);
            using (Brush fill = new SolidBrush(ParseColor(shape.FillColorHex, Color.FromArgb(91, 140, 255))))
            using (Pen line = new Pen(ParseColor(shape.LineColorHex, Color.FromArgb(53, 106, 230)),
                Math.Max(1f, shape.LineWidthPoints)))
            {
                GraphicsPath path = CreateShapePath(shape.Kind, rect);
                if (path != null)
                {
                    graphics.FillPath(fill, path);
                    graphics.DrawPath(line, path);
                    path.Dispose();
                }
            }

            DrawSelectionIfNeeded(graphics, EditorObjectKind.Shape, index, rect);
        }

        private void DrawImage(Graphics graphics, PresentationImage item, int index)
        {
            RectangleF rect = ToRectangle(item.X, item.Y, item.Width, item.Height);

            if (item.Data != null && item.Data.Length > 0)
            {
                try
                {
                    using (MemoryStream stream = new MemoryStream(item.Data, false))
                    using (Image source = Image.FromStream(stream))
                    using (Bitmap bitmap = new Bitmap(source))
                    {
                        graphics.DrawImage(bitmap, Rectangle.Round(rect));
                    }
                }
                catch
                {
                    using (Brush placeholder = new SolidBrush(Color.FromArgb(230, 232, 236)))
                        graphics.FillRectangle(placeholder, rect);
                }
            }

            DrawSelectionIfNeeded(graphics, EditorObjectKind.Image, index, rect);
        }

        private void DrawSelectionIfNeeded(
            Graphics graphics,
            EditorObjectKind kind,
            int index,
            RectangleF rect)
        {
            if (SelectedObjectKind != kind || SelectedObjectIndex != index)
                return;

            using (Pen pen = new Pen(ApplicationTheme.Accent, 2f))
                graphics.DrawRectangle(pen, Rectangle.Round(rect));

            Rectangle handle = ResizeHandle(rect);
            using (Brush brush = new SolidBrush(ApplicationTheme.Accent))
                graphics.FillRectangle(brush, handle);
        }

        private bool HitTest(
            Point point,
            out EditorObjectKind kind,
            out int index,
            out RectangleF rect)
        {
            kind = EditorObjectKind.None;
            index = -1;
            rect = RectangleF.Empty;
            PresentationSlide slide = GetSlide();
            if (slide == null)
                return false;

            for (int i = slide.TextBoxes.Count - 1; i >= 0; i--)
            {
                RectangleF candidate = ToRectangle(
                    slide.TextBoxes[i].X,
                    slide.TextBoxes[i].Y,
                    slide.TextBoxes[i].Width,
                    slide.TextBoxes[i].Height);
                if (candidate.Contains(point))
                {
                    kind = EditorObjectKind.TextBox;
                    index = i;
                    rect = candidate;
                    return true;
                }
            }

            for (int i = slide.Images.Count - 1; i >= 0; i--)
            {
                RectangleF candidate = ToRectangle(
                    slide.Images[i].X,
                    slide.Images[i].Y,
                    slide.Images[i].Width,
                    slide.Images[i].Height);
                if (candidate.Contains(point))
                {
                    kind = EditorObjectKind.Image;
                    index = i;
                    rect = candidate;
                    return true;
                }
            }

            for (int i = slide.Shapes.Count - 1; i >= 0; i--)
            {
                RectangleF candidate = ToRectangle(
                    slide.Shapes[i].X,
                    slide.Shapes[i].Y,
                    slide.Shapes[i].Width,
                    slide.Shapes[i].Height);
                if (candidate.Contains(point))
                {
                    kind = EditorObjectKind.Shape;
                    index = i;
                    rect = candidate;
                    return true;
                }
            }

            return false;
        }

        private RectangleF GetSelectedRectangle()
        {
            long x;
            long y;
            long width;
            long height;
            if (!GetSelectedGeometry(out x, out y, out width, out height))
                return RectangleF.Empty;
            return ToRectangle(x, y, width, height);
        }

        private bool GetSelectedGeometry(
            out long x,
            out long y,
            out long width,
            out long height)
        {
            x = 0;
            y = 0;
            width = 1;
            height = 1;
            PresentationTextBox text = GetSelectedTextBox();
            if (text != null)
            {
                x = text.X; y = text.Y; width = text.Width; height = text.Height;
                return true;
            }

            PresentationShape shape = GetSelectedShape();
            if (shape != null)
            {
                x = shape.X; y = shape.Y; width = shape.Width; height = shape.Height;
                return true;
            }

            PresentationImage image = GetSelectedImage();
            if (image != null)
            {
                x = image.X; y = image.Y; width = image.Width; height = image.Height;
                return true;
            }

            return false;
        }

        private void SetSelectedGeometry(long x, long y, long width, long height)
        {
            PresentationTextBox text = GetSelectedTextBox();
            if (text != null)
            {
                text.X = x; text.Y = y; text.Width = width; text.Height = height;
                return;
            }

            PresentationShape shape = GetSelectedShape();
            if (shape != null)
            {
                shape.X = x; shape.Y = y; shape.Width = width; shape.Height = height;
                return;
            }

            PresentationImage image = GetSelectedImage();
            if (image != null)
            {
                image.X = x; image.Y = y; image.Width = width; image.Height = height;
            }
        }

        private long PixelsToEmuX(int pixels)
        {
            if (slideRectangle.Width <= 0)
                return 0;
            return (long)Math.Round(pixels * (Document.WidthEmu / (double)slideRectangle.Width));
        }

        private long PixelsToEmuY(int pixels)
        {
            if (slideRectangle.Height <= 0)
                return 0;
            return (long)Math.Round(pixels * (Document.HeightEmu / (double)slideRectangle.Height));
        }

        private static Rectangle ResizeHandle(RectangleF rect)
        {
            const int size = 10;
            return new Rectangle(
                (int)Math.Round(rect.Right) - size / 2,
                (int)Math.Round(rect.Bottom) - size / 2,
                size,
                size);
        }

        private static GraphicsPath CreateShapePath(
            PresentationShapeKind kind,
            RectangleF rect)
        {
            GraphicsPath path = new GraphicsPath();

            if (kind == PresentationShapeKind.Ellipse)
            {
                path.AddEllipse(rect);
                return path;
            }

            if (kind == PresentationShapeKind.Triangle)
            {
                path.AddPolygon(new PointF[]
                {
                    new PointF(rect.Left + rect.Width / 2f, rect.Top),
                    new PointF(rect.Right, rect.Bottom),
                    new PointF(rect.Left, rect.Bottom)
                });
                return path;
            }

            if (kind == PresentationShapeKind.Diamond)
            {
                path.AddPolygon(new PointF[]
                {
                    new PointF(rect.Left + rect.Width / 2f, rect.Top),
                    new PointF(rect.Right, rect.Top + rect.Height / 2f),
                    new PointF(rect.Left + rect.Width / 2f, rect.Bottom),
                    new PointF(rect.Left, rect.Top + rect.Height / 2f)
                });
                return path;
            }

            if (kind == PresentationShapeKind.RoundedRectangle)
            {
                float radius = Math.Max(4f, Math.Min(rect.Width, rect.Height) * 0.12f);
                float diameter = radius * 2f;
                path.AddArc(rect.Left, rect.Top, diameter, diameter, 180f, 90f);
                path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270f, 90f);
                path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0f, 90f);
                path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90f, 90f);
                path.CloseFigure();
                return path;
            }

            path.AddRectangle(rect);
            return path;
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

        private static Color ParseColor(string value, Color fallback)
        {
            string candidate = (value ?? string.Empty).Trim().TrimStart('#');
            int parsed;
            if (candidate.Length == 6 &&
                int.TryParse(candidate, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out parsed))
            {
                return Color.FromArgb(
                    (parsed >> 16) & 0xFF,
                    (parsed >> 8) & 0xFF,
                    parsed & 0xFF);
            }

            return fallback;
        }
    }
}
