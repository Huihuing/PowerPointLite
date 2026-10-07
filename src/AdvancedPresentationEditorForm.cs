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

    internal sealed class EditorSelectionEntry
    {
        public EditorObjectKind Kind { get; set; }
        public int Index { get; set; }

        public EditorSelectionEntry Clone()
        {
            EditorSelectionEntry copy =
                new EditorSelectionEntry();
            copy.Kind = Kind;
            copy.Index = Index;
            return copy;
        }
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
        private readonly Button sendBackButton;
        private readonly Button bringFrontButton;

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
            mainToolbar.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (Pen divider =
                    new Pen(
                        ApplicationTheme.Divider))
                {
                    int bottom =
                        Math.Max(
                            0,
                            mainToolbar.ClientSize.Height -
                            1);

                    e.Graphics.DrawLine(
                        divider,
                        0,
                        bottom,
                        Math.Max(
                            0,
                            mainToolbar.ClientSize.Width -
                            1),
                        bottom);

                    int[] separators =
                        new int[]
                        {
                            161,
                            306,
                            609,
                            754,
                            948
                        };

                    for (int i = 0;
                         i < separators.Length;
                         i++)
                    {
                        e.Graphics.DrawLine(
                            divider,
                            separators[i],
                            12,
                            separators[i],
                            Math.Max(
                                12,
                                mainToolbar.ClientSize.Height -
                                12));
                    }
                }
            };
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
            sendBackButton = MakeButton("Send Back", 760, 84);
            bringFrontButton = MakeButton("Bring Front", 850, 92);

            saveButton.Tag = "Primary";
            ApplicationTheme.ApplyButton(
                saveButton);
            deleteObjectButton.Tag = "Danger";
            ApplicationTheme.ApplyButton(
                deleteObjectButton);

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
            mainToolbar.Controls.Add(sendBackButton);
            mainToolbar.Controls.Add(bringFrontButton);

            Label quickShapeLabel =
                MakeLabel(
                    "Shape type",
                    0,
                    0,
                    68,
                    28,
                    false);
            quickShapeLabel.ForeColor =
                ApplicationTheme.SecondaryText;
            quickShapeLabel.TextAlign =
                ContentAlignment.MiddleRight;
            mainToolbar.Controls.Add(
                quickShapeLabel);

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
                    int labelGap = 8;
                    int selectorLabelWidth = 68;
                    int minimumSelectorWidth = 104;
                    int desiredSelectorWidth = 148;
                    int selectorRight =
                        Math.Max(
                            bringFrontButton.Right +
                            gap,
                            mainToolbar.ClientSize.Width -
                            10);

                    int availableSelectorWidth =
                        selectorRight -
                        bringFrontButton.Right -
                        gap;

                    if (availableSelectorWidth <
                        minimumSelectorWidth +
                        selectorLabelWidth +
                        labelGap)
                    {
                        quickShapeLabel.Visible = false;
                        quickShapeKind.Visible = false;
                        return;
                    }

                    quickShapeLabel.Visible = true;
                    quickShapeKind.Visible = true;
                    quickShapeKind.Width =
                        Math.Min(
                            desiredSelectorWidth,
                            availableSelectorWidth -
                            selectorLabelWidth -
                            labelGap);
                    quickShapeKind.Left =
                        selectorRight -
                        quickShapeKind.Width;
                    quickShapeLabel.Left =
                        quickShapeKind.Left -
                        labelGap -
                        selectorLabelWidth;
                    quickShapeLabel.Top = 10;
                    quickShapeLabel.Width =
                        selectorLabelWidth;
                    quickShapeLabel.Height =
                        quickShapeKind.Height;
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
            sendBackButton.Click += delegate { MoveSelectedObjectLayer(false); };
            bringFrontButton.Click += delegate { MoveSelectedObjectLayer(true); };

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
            properties.AutoScroll = true;
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
            title.Font =
                new Font(
                    Font.FontFamily,
                    11f,
                    FontStyle.Bold);
            properties.Controls.Add(title);

            objectTypeLabel = MakeLabel("No object selected", 0, 30, 260, 28, false);
            objectTypeLabel.ForeColor = ApplicationTheme.SecondaryText;
            objectTypeLabel.BackColor = ApplicationTheme.Surface;
            objectTypeLabel.Padding = new Padding(8, 0, 8, 0);
            objectTypeLabel.TextAlign = ContentAlignment.MiddleLeft;
            properties.Controls.Add(objectTypeLabel);

            Label textLabel = MakeLabel("Text", 0, 68, 260, 20, true);
            textLabel.ForeColor = ApplicationTheme.Accent;

            properties.Controls.Add(textLabel);

            textEditor = new TextBox();
            textEditor.Left = 0;
            textEditor.Top = 92;
            textEditor.Width = 270;
            textEditor.Height = 108;
            textEditor.Multiline = true;
            textEditor.ScrollBars = ScrollBars.Vertical;
            textEditor.BackColor = ApplicationTheme.Surface;
            textEditor.ForeColor = ApplicationTheme.PrimaryText;
            textEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(textEditor);

            Label fontLabel = MakeLabel("Font", 0, 210, 260, 20, false);
            properties.Controls.Add(fontLabel);

            fontPicker = new ComboBox();
            fontPicker.Left = 0;
            fontPicker.Top = 232;
            fontPicker.Width = 270;
            fontPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            fontPicker.FlatStyle = FlatStyle.Flat;
            fontPicker.BackColor = ApplicationTheme.Surface;
            fontPicker.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(fontPicker);

            List<string> fonts = SystemFontCatalog.GetInstalledFamilyNames();
            for (int i = 0; i < fonts.Count; i++)
                fontPicker.Items.Add(fonts[i]);

            Label sizeLabel = MakeLabel("Size", 0, 264, 70, 20, false);
            properties.Controls.Add(sizeLabel);

            fontSize = new NumericUpDown();
            fontSize.Left = 0;
            fontSize.Top = 286;
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
            boldCheck.Top = 288;
            boldCheck.Width = 68;
            boldCheck.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(boldCheck);

            italicCheck = new CheckBox();
            italicCheck.Text = "Italic";
            italicCheck.Left = 172;
            italicCheck.Top = 288;
            italicCheck.Width = 72;
            italicCheck.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(italicCheck);

            Label alignLabel = MakeLabel("Alignment", 0, 322, 100, 20, false);
            properties.Controls.Add(alignLabel);

            alignmentPicker = new ComboBox();
            alignmentPicker.Left = 0;
            alignmentPicker.Top = 344;
            alignmentPicker.Width = 126;
            alignmentPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            alignmentPicker.FlatStyle = FlatStyle.Flat;
            alignmentPicker.BackColor = ApplicationTheme.Surface;
            alignmentPicker.ForeColor = ApplicationTheme.PrimaryText;
            alignmentPicker.Items.Add("Left");
            alignmentPicker.Items.Add("Center");
            alignmentPicker.Items.Add("Right");
            properties.Controls.Add(alignmentPicker);

            Label textColorLabel = MakeLabel("Text color", 142, 322, 100, 20, false);
            properties.Controls.Add(textColorLabel);

            textColorEditor = new TextBox();
            textColorEditor.Left = 142;
            textColorEditor.Top = 344;
            textColorEditor.Width = 128;
            textColorEditor.MaxLength = 7;
            textColorEditor.BackColor = ApplicationTheme.Surface;
            textColorEditor.ForeColor = ApplicationTheme.PrimaryText;
            textColorEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(textColorEditor);

            Label shapeLabel = MakeLabel("Shape", 0, 390, 100, 20, true);
            shapeLabel.ForeColor = ApplicationTheme.Accent;
            properties.Controls.Add(shapeLabel);

            shapeKindPicker = new ComboBox();
            shapeKindPicker.Left = 0;
            shapeKindPicker.Top = 414;
            shapeKindPicker.Width = 270;
            shapeKindPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            shapeKindPicker.FlatStyle = FlatStyle.Flat;
            shapeKindPicker.BackColor = ApplicationTheme.Surface;
            shapeKindPicker.ForeColor = ApplicationTheme.PrimaryText;
            AddShapeKindItems(shapeKindPicker);
            properties.Controls.Add(shapeKindPicker);

            Label fillLabel = MakeLabel("Fill", 0, 448, 80, 20, false);
            properties.Controls.Add(fillLabel);
            fillColorEditor = new TextBox();
            fillColorEditor.Left = 0;
            fillColorEditor.Top = 470;
            fillColorEditor.Width = 126;
            fillColorEditor.MaxLength = 7;
            fillColorEditor.BackColor = ApplicationTheme.Surface;
            fillColorEditor.ForeColor = ApplicationTheme.PrimaryText;
            fillColorEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(fillColorEditor);

            Label lineLabel = MakeLabel("Line", 142, 448, 80, 20, false);
            properties.Controls.Add(lineLabel);
            lineColorEditor = new TextBox();
            lineColorEditor.Left = 142;
            lineColorEditor.Top = 470;
            lineColorEditor.Width = 128;
            lineColorEditor.MaxLength = 7;
            lineColorEditor.BackColor = ApplicationTheme.Surface;
            lineColorEditor.ForeColor = ApplicationTheme.PrimaryText;
            lineColorEditor.BorderStyle = BorderStyle.FixedSingle;
            properties.Controls.Add(lineColorEditor);

            Label hint = MakeLabel(
                "Drag objects to move. Drag the lower-right handle to resize. Arrow keys nudge. Ctrl+Shift+Up/Down changes layer order.",
                0,
                516,
                270,
                80,
                false);
            hint.ForeColor = ApplicationTheme.SecondaryText;
            properties.Controls.Add(hint);

            Action layoutPropertyFields =
                delegate
                {
                    const int contentLeft = 14;
                    const int contentRight = 14;
                    const int columnGap = 14;

                    int contentWidth =
                        Math.Max(
                            220,
                            properties.ClientSize.Width -
                            contentLeft -
                            contentRight);

                    int halfWidth =
                        Math.Max(
                            96,
                            (contentWidth -
                             columnGap) /
                            2);

                    int rightColumnLeft =
                        contentLeft +
                        halfWidth +
                        columnGap;

                    title.Left =
                        contentLeft;
                    title.Width =
                        contentWidth;

                    objectTypeLabel.Left =
                        contentLeft;
                    objectTypeLabel.Width =
                        contentWidth;

                    textLabel.Left =
                        contentLeft;
                    textLabel.Width =
                        contentWidth;

                    textEditor.Left =
                        contentLeft;
                    textEditor.Width =
                        contentWidth;

                    fontLabel.Left =
                        contentLeft;
                    fontLabel.Width =
                        contentWidth;

                    fontPicker.Left =
                        contentLeft;
                    fontPicker.Width =
                        contentWidth;

                    sizeLabel.Left =
                        contentLeft;
                    fontSize.Left =
                        contentLeft;
                    boldCheck.Left =
                        contentLeft +
                        92;
                    italicCheck.Left =
                        contentLeft +
                        164;

                    alignLabel.Left =
                        contentLeft;
                    alignLabel.Width =
                        halfWidth;

                    alignmentPicker.Left =
                        contentLeft;
                    alignmentPicker.Width =
                        halfWidth;

                    textColorLabel.Left =
                        rightColumnLeft;
                    textColorLabel.Width =
                        halfWidth;

                    textColorEditor.Left =
                        rightColumnLeft;
                    textColorEditor.Width =
                        halfWidth;

                    shapeLabel.Left =
                        contentLeft;
                    shapeLabel.Width =
                        contentWidth;

                    shapeKindPicker.Left =
                        contentLeft;
                    shapeKindPicker.Width =
                        contentWidth;

                    fillLabel.Left =
                        contentLeft;
                    fillLabel.Width =
                        halfWidth;

                    fillColorEditor.Left =
                        contentLeft;
                    fillColorEditor.Width =
                        halfWidth;

                    lineLabel.Left =
                        rightColumnLeft;
                    lineLabel.Width =
                        halfWidth;

                    lineColorEditor.Left =
                        rightColumnLeft;
                    lineColorEditor.Width =
                        halfWidth;

                    hint.Left =
                        contentLeft;
                    hint.Width =
                        contentWidth;

                    properties.AutoScrollMinSize =
                        new Size(
                            0,
                            hint.Bottom +
                            16);
                };

            properties.Resize += delegate
            {
                layoutPropertyFields();
            };

            layoutPropertyFields();

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

            Action layoutEditorColumns =
                delegate
                {
                    int width =
                        Math.Max(
                            1,
                            ClientSize.Width);

                    slidePanel.Width =
                        Math.Max(
                            184,
                            Math.Min(
                                220,
                                width /
                                6));

                    properties.Width =
                        Math.Max(
                            286,
                            Math.Min(
                                300,
                                width /
                                4));

                    int canvasPadding =
                        width < 1220
                            ? 14
                            : 22;

                    canvasHost.Padding =
                        new Padding(
                            canvasPadding);

                    slidePanel.Invalidate();
                    properties.Invalidate();
                };

            Resize += delegate
            {
                layoutEditorColumns();
            };

            layoutEditorColumns();

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
                UpdateStatus();
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

            if (keyData ==
                (Keys.Control |
                 Keys.Shift |
                 Keys.Up))
            {
                MoveSelectedObjectLayer(true);
                return true;
            }

            if (keyData ==
                (Keys.Control |
                 Keys.Shift |
                 Keys.Down))
            {
                MoveSelectedObjectLayer(false);
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
            int slideIndex =
                canvas.SelectedSlideIndex;
            List<EditorSelectionEntry> selected =
                canvas.GetSelectedObjects();

            if (slideIndex < 0 ||
                selected.Count == 0)
            {
                return;
            }

            List<int> textIndexes =
                new List<int>();
            List<int> shapeIndexes =
                new List<int>();
            List<int> imageIndexes =
                new List<int>();

            for (int i = 0;
                 i < selected.Count;
                 i++)
            {
                EditorSelectionEntry entry =
                    selected[i];

                if (entry.Kind ==
                    EditorObjectKind.TextBox)
                {
                    textIndexes.Add(
                        entry.Index);
                }
                else if (entry.Kind ==
                         EditorObjectKind.Shape)
                {
                    shapeIndexes.Add(
                        entry.Index);
                }
                else if (entry.Kind ==
                         EditorObjectKind.Image)
                {
                    imageIndexes.Add(
                        entry.Index);
                }
            }

            textIndexes.Sort();
            shapeIndexes.Sort();
            imageIndexes.Sort();

            CaptureHistory();
            bool removed = false;

            for (int i =
                     textIndexes.Count - 1;
                 i >= 0;
                 i--)
            {
                removed =
                    session.RemoveTextBox(
                        slideIndex,
                        textIndexes[i]) ||
                    removed;
            }

            for (int i =
                     shapeIndexes.Count - 1;
                 i >= 0;
                 i--)
            {
                removed =
                    session.RemoveShape(
                        slideIndex,
                        shapeIndexes[i]) ||
                    removed;
            }

            for (int i =
                     imageIndexes.Count - 1;
                 i >= 0;
                 i--)
            {
                removed =
                    session.RemoveImage(
                        slideIndex,
                        imageIndexes[i]) ||
                    removed;
            }

            if (!removed)
                return;

            canvas.Document = session.Document;
            canvas.ClearSelection();
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private bool TryGetSelectedLayerKind(
            out PresentationLayerKind kind)
        {
            kind =
                PresentationLayerKind.TextBox;

            if (canvas.SelectedObjectKind ==
                EditorObjectKind.TextBox)
            {
                kind =
                    PresentationLayerKind.TextBox;
                return true;
            }

            if (canvas.SelectedObjectKind ==
                EditorObjectKind.Shape)
            {
                kind =
                    PresentationLayerKind.Shape;
                return true;
            }

            if (canvas.SelectedObjectKind ==
                EditorObjectKind.Image)
            {
                kind =
                    PresentationLayerKind.Image;
                return true;
            }

            return false;
        }

        private void MoveSelectedObjectLayer(
            bool toFront)
        {
            PresentationSlide slide =
                GetSelectedSlide();

            PresentationLayerKind kind;
            int index =
                canvas.SelectedObjectIndex;

            if (slide == null ||
                canvas.SelectedObjectCount != 1 ||
                index < 0 ||
                !TryGetSelectedLayerKind(
                    out kind))
            {
                return;
            }

            bool canMove =
                toFront
                    ? slide.CanMoveObjectToFront(
                        kind,
                        index)
                    : slide.CanMoveObjectToBack(
                        kind,
                        index);

            if (!canMove)
                return;

            CaptureHistory();

            bool moved =
                toFront
                    ? slide.MoveObjectToFront(
                        kind,
                        index)
                    : slide.MoveObjectToBack(
                        kind,
                        index);

            if (!moved)
                return;

            propertyEditSnapshotActive = false;
            session.MarkDirty();
            canvas.Invalidate();
            UpdateStatus();
        }

        private void CopySelectedObject()
        {
            copiedTextBox = null;
            copiedShape = null;
            copiedImage = null;

            PresentationSlide slide = GetSelectedSlide();
            if (slide == null ||
                canvas.SelectedObjectCount != 1 ||
                canvas.SelectedObjectIndex < 0)
            {
                return;
            }

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
                slide.RegisterObjectOrder(
                    PresentationLayerKind.TextBox,
                    slide.TextBoxes.Count - 1);
                canvas.SelectObject(EditorObjectKind.TextBox, slide.TextBoxes.Count - 1);
            }
            else if (copiedShape != null)
            {
                PresentationShape copy = copiedShape.Clone();
                copy.X += offset;
                copy.Y += offset;
                slide.Shapes.Add(copy);
                slide.RegisterObjectOrder(
                    PresentationLayerKind.Shape,
                    slide.Shapes.Count - 1);
                canvas.SelectObject(EditorObjectKind.Shape, slide.Shapes.Count - 1);
            }
            else if (copiedImage != null)
            {
                PresentationImage copy = copiedImage.Clone();
                copy.X += offset;
                copy.Y += offset;
                slide.Images.Add(copy);
                slide.RegisterObjectOrder(
                    PresentationLayerKind.Image,
                    slide.Images.Count - 1);
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
                int selectedCount =
                    canvas.SelectedObjectCount;
                PresentationTextBox text = canvas.GetSelectedTextBox();
                PresentationShape shape = canvas.GetSelectedShape();
                PresentationImage image = canvas.GetSelectedImage();

                bool textEnabled = text != null;
                bool shapeEnabled = shape != null;

                objectTypeLabel.Text =
                    selectedCount > 1
                        ? selectedCount.ToString() +
                          " objects selected"
                        : text != null
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

            int selectedCount =
                canvas == null
                    ? 0
                    : canvas.SelectedObjectCount;

            statusLabel.Text =
                (session.IsDirty ? "Modified  •  " : "Saved  •  ") +
                fileName + "  •  " +
                session.Document.Slides.Count.ToString() +
                " slide(s)" +
                (selectedCount > 0
                    ? "  •  " +
                      selectedCount.ToString() +
                      (selectedCount == 1
                          ? " object selected"
                          : " objects selected")
                    : string.Empty);
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            undoButton.Enabled = history.CanUndo;
            redoButton.Enabled = history.CanRedo;
            bool hasSelection =
                canvas != null &&
                canvas.SelectedObjectCount > 0;
            bool singleSelection =
                hasSelection &&
                canvas.SelectedObjectCount == 1;
            deleteObjectButton.Enabled = hasSelection;
            copyButton.Enabled = singleSelection;
            pasteButton.Enabled = copiedTextBox != null || copiedShape != null || copiedImage != null;

            PresentationSlide slide =
                GetSelectedSlide();
            PresentationLayerKind layerKind =
                PresentationLayerKind.TextBox;
            bool hasLayerSelection =
                singleSelection &&
                slide != null &&
                TryGetSelectedLayerKind(
                    out layerKind);

            sendBackButton.Enabled =
                hasLayerSelection &&
                slide.CanMoveObjectToBack(
                    layerKind,
                    canvas.SelectedObjectIndex);
            bringFrontButton.Enabled =
                hasLayerSelection &&
                slide.CanMoveObjectToFront(
                    layerKind,
                    canvas.SelectedObjectIndex);
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
        private sealed class SelectionGeometrySnapshot
        {
            public EditorObjectKind Kind;
            public int Index;
            public long X;
            public long Y;
            public long Width;
            public long Height;
        }

        private readonly List<EditorSelectionEntry> selectedObjects =
            new List<EditorSelectionEntry>();
        private readonly List<SelectionGeometrySnapshot> dragSelection =
            new List<SelectionGeometrySnapshot>();
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
        public int SelectedObjectCount
        {
            get { return selectedObjects.Count; }
        }

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
            bool changed =
                selectedObjects.Count > 0 ||
                SelectedObjectKind !=
                    EditorObjectKind.None;

            selectedObjects.Clear();
            SelectedObjectKind =
                EditorObjectKind.None;
            SelectedObjectIndex = -1;
            Invalidate();

            if (changed &&
                SelectionChanged != null)
            {
                SelectionChanged(
                    this,
                    EventArgs.Empty);
            }
        }

        public void SelectObject(
            EditorObjectKind kind,
            int index)
        {
            selectedObjects.Clear();
            AddSelectionEntry(
                kind,
                index);
            SetPrimarySelection(
                kind,
                index);
            NotifySelectionChanged();
        }

        public List<EditorSelectionEntry>
            GetSelectedObjects()
        {
            List<EditorSelectionEntry> result =
                new List<EditorSelectionEntry>();

            for (int i = 0;
                 i < selectedObjects.Count;
                 i++)
            {
                result.Add(
                    selectedObjects[i].Clone());
            }

            return result;
        }

        public bool IsObjectSelected(
            EditorObjectKind kind,
            int index)
        {
            return FindSelectionIndex(
                kind,
                index) >= 0;
        }

        private void ToggleObjectSelection(
            EditorObjectKind kind,
            int index)
        {
            int existing =
                FindSelectionIndex(
                    kind,
                    index);

            if (existing >= 0)
            {
                selectedObjects.RemoveAt(
                    existing);

                if (SelectedObjectKind == kind &&
                    SelectedObjectIndex == index)
                {
                    if (selectedObjects.Count > 0)
                    {
                        EditorSelectionEntry replacement =
                            selectedObjects[
                                selectedObjects.Count - 1];
                        SetPrimarySelection(
                            replacement.Kind,
                            replacement.Index);
                    }
                    else
                    {
                        SetPrimarySelection(
                            EditorObjectKind.None,
                            -1);
                    }
                }
            }
            else
            {
                AddSelectionEntry(
                    kind,
                    index);
                SetPrimarySelection(
                    kind,
                    index);
            }

            NotifySelectionChanged();
        }

        private void PromotePrimarySelection(
            EditorObjectKind kind,
            int index)
        {
            if (!IsObjectSelected(
                    kind,
                    index))
            {
                return;
            }

            if (SelectedObjectKind == kind &&
                SelectedObjectIndex == index)
            {
                return;
            }

            SetPrimarySelection(
                kind,
                index);
            NotifySelectionChanged();
        }

        private void AddSelectionEntry(
            EditorObjectKind kind,
            int index)
        {
            if (kind == EditorObjectKind.None ||
                index < 0 ||
                IsObjectSelected(
                    kind,
                    index))
            {
                return;
            }

            EditorSelectionEntry entry =
                new EditorSelectionEntry();
            entry.Kind = kind;
            entry.Index = index;
            selectedObjects.Add(entry);
        }

        private int FindSelectionIndex(
            EditorObjectKind kind,
            int index)
        {
            for (int i = 0;
                 i < selectedObjects.Count;
                 i++)
            {
                EditorSelectionEntry entry =
                    selectedObjects[i];

                if (entry.Kind == kind &&
                    entry.Index == index)
                {
                    return i;
                }
            }

            return -1;
        }

        private void SetPrimarySelection(
            EditorObjectKind kind,
            int index)
        {
            SelectedObjectKind = kind;
            SelectedObjectIndex = index;
        }

        private void NotifySelectionChanged()
        {
            Invalidate();

            if (SelectionChanged != null)
                SelectionChanged(
                    this,
                    EventArgs.Empty);
        }

        public PresentationTextBox GetSelectedTextBox()
        {
            PresentationSlide slide = GetSlide();
            if (slide == null ||
                selectedObjects.Count != 1 ||
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
                selectedObjects.Count != 1 ||
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
                selectedObjects.Count != 1 ||
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
                    DrawShape(
                        e.Graphics,
                        slide.Shapes[entry.Index],
                        entry.Index);
                }
                else if (entry.Kind ==
                             PresentationLayerKind.Image &&
                         entry.Index >= 0 &&
                         entry.Index < slide.Images.Count)
                {
                    DrawImage(
                        e.Graphics,
                        slide.Images[entry.Index],
                        entry.Index);
                }
                else if (entry.Kind ==
                             PresentationLayerKind.TextBox &&
                         entry.Index >= 0 &&
                         entry.Index < slide.TextBoxes.Count)
                {
                    DrawTextBox(
                        e.Graphics,
                        slide.TextBoxes[entry.Index],
                        entry.Index);
                }
            }

            if (selectedObjects.Count > 1)
            {
                DrawMultiSelectionBounds(
                    e.Graphics);
            }
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

            bool controlPressed =
                (ModifierKeys & Keys.Control) ==
                Keys.Control;

            if (!HitTest(e.Location, out kind, out index, out rect))
            {
                if (!controlPressed)
                    ClearSelection();
                return;
            }

            if (controlPressed)
            {
                ToggleObjectSelection(
                    kind,
                    index);

                if (!IsObjectSelected(
                        kind,
                        index))
                {
                    return;
                }
            }
            else if (selectedObjects.Count > 1 &&
                     IsObjectSelected(
                         kind,
                         index))
            {
                PromotePrimarySelection(
                    kind,
                    index);
            }
            else
            {
                SelectObject(
                    kind,
                    index);
            }

            dragging = true;
            resizing =
                selectedObjects.Count == 1 &&
                ResizeHandle(rect).Contains(
                    e.Location);
            historyStarted = false;
            dragStart = e.Location;
            CaptureDragSelection();
            GetSelectedGeometry(out originalX, out originalY, out originalWidth, out originalHeight);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (!dragging)
            {
                RectangleF selected =
                    selectedObjects.Count == 1
                        ? GetSelectedRectangle()
                        : RectangleF.Empty;
                Cursor =
                    selected != RectangleF.Empty &&
                    ResizeHandle(selected).Contains(
                        e.Location)
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
                MoveDragSelection(
                    dxEmu,
                    dyEmu);
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

            if (code == Keys.Delete && selectedObjects.Count > 0)
            {
                if (DeleteRequested != null)
                    DeleteRequested(this, EventArgs.Empty);
                return true;
            }

            if (code == Keys.Left || code == Keys.Right ||
                code == Keys.Up || code == Keys.Down)
            {
                if (selectedObjects.Count == 0)
                    return base.ProcessCmdKey(ref msg, keyData);

                if (TransformStarting != null)
                    TransformStarting(this, EventArgs.Empty);

                long step = shift ? 228600L : 45720L;
                long dx = 0L;
                long dy = 0L;

                if (code == Keys.Left)
                    dx = -step;
                else if (code == Keys.Right)
                    dx = step;
                else if (code == Keys.Up)
                    dy = -step;
                else if (code == Keys.Down)
                    dy = step;

                MoveCurrentSelection(
                    dx,
                    dy);
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
            if (!IsObjectSelected(
                    kind,
                    index))
            {
                return;
            }

            bool primary =
                SelectedObjectKind == kind &&
                SelectedObjectIndex == index;

            using (Pen pen = new Pen(
                ApplicationTheme.Accent,
                primary ? 2f : 1.4f))
            {
                if (selectedObjects.Count > 1 &&
                    !primary)
                {
                    pen.DashStyle =
                        DashStyle.Dash;
                }

                graphics.DrawRectangle(
                    pen,
                    Rectangle.Round(rect));
            }

            if (selectedObjects.Count == 1 &&
                primary)
            {
                Rectangle handle =
                    ResizeHandle(rect);
                using (Brush brush =
                    new SolidBrush(
                        ApplicationTheme.Accent))
                {
                    graphics.FillRectangle(
                        brush,
                        handle);
                }
            }
        }


        private void DrawMultiSelectionBounds(
            Graphics graphics)
        {
            RectangleF bounds =
                GetSelectionBounds();

            if (bounds == RectangleF.Empty)
                return;

            bounds.Inflate(
                4f,
                4f);

            using (Pen pen = new Pen(
                ApplicationTheme.Accent,
                1f))
            {
                pen.DashStyle =
                    DashStyle.Dash;

                graphics.DrawRectangle(
                    pen,
                    Rectangle.Round(bounds));
            }
        }

        private RectangleF GetSelectionBounds()
        {
            RectangleF bounds =
                RectangleF.Empty;

            for (int i = 0;
                 i < selectedObjects.Count;
                 i++)
            {
                EditorSelectionEntry entry =
                    selectedObjects[i];
                RectangleF item =
                    GetObjectRectangle(
                        entry.Kind,
                        entry.Index);

                if (item == RectangleF.Empty)
                    continue;

                bounds =
                    bounds == RectangleF.Empty
                        ? item
                        : RectangleF.Union(
                            bounds,
                            item);
            }

            return bounds;
        }

        private RectangleF GetObjectRectangle(
            EditorObjectKind kind,
            int index)
        {
            long x;
            long y;
            long width;
            long height;

            if (!GetObjectGeometry(
                    kind,
                    index,
                    out x,
                    out y,
                    out width,
                    out height))
            {
                return RectangleF.Empty;
            }

            return ToRectangle(
                x,
                y,
                width,
                height);
        }

        private void CaptureDragSelection()
        {
            dragSelection.Clear();

            for (int i = 0;
                 i < selectedObjects.Count;
                 i++)
            {
                EditorSelectionEntry entry =
                    selectedObjects[i];
                long x;
                long y;
                long width;
                long height;

                if (!GetObjectGeometry(
                        entry.Kind,
                        entry.Index,
                        out x,
                        out y,
                        out width,
                        out height))
                {
                    continue;
                }

                SelectionGeometrySnapshot item =
                    new SelectionGeometrySnapshot();
                item.Kind = entry.Kind;
                item.Index = entry.Index;
                item.X = x;
                item.Y = y;
                item.Width = width;
                item.Height = height;
                dragSelection.Add(item);
            }
        }

        private void MoveDragSelection(
            long dx,
            long dy)
        {
            MoveGeometrySnapshots(
                dragSelection,
                dx,
                dy);
        }

        private void MoveCurrentSelection(
            long dx,
            long dy)
        {
            List<SelectionGeometrySnapshot> snapshot =
                new List<SelectionGeometrySnapshot>();

            for (int i = 0;
                 i < selectedObjects.Count;
                 i++)
            {
                EditorSelectionEntry entry =
                    selectedObjects[i];
                long x;
                long y;
                long width;
                long height;

                if (!GetObjectGeometry(
                        entry.Kind,
                        entry.Index,
                        out x,
                        out y,
                        out width,
                        out height))
                {
                    continue;
                }

                SelectionGeometrySnapshot item =
                    new SelectionGeometrySnapshot();
                item.Kind = entry.Kind;
                item.Index = entry.Index;
                item.X = x;
                item.Y = y;
                item.Width = width;
                item.Height = height;
                snapshot.Add(item);
            }

            MoveGeometrySnapshots(
                snapshot,
                dx,
                dy);
        }

        private void MoveGeometrySnapshots(
            List<SelectionGeometrySnapshot> snapshot,
            long dx,
            long dy)
        {
            if (snapshot == null ||
                snapshot.Count == 0 ||
                Document == null)
            {
                return;
            }

            long minX =
                long.MaxValue;
            long minY =
                long.MaxValue;
            long maxRight =
                long.MinValue;
            long maxBottom =
                long.MinValue;

            for (int i = 0;
                 i < snapshot.Count;
                 i++)
            {
                SelectionGeometrySnapshot item =
                    snapshot[i];

                minX =
                    Math.Min(
                        minX,
                        item.X);
                minY =
                    Math.Min(
                        minY,
                        item.Y);
                maxRight =
                    Math.Max(
                        maxRight,
                        item.X +
                        item.Width);
                maxBottom =
                    Math.Max(
                        maxBottom,
                        item.Y +
                        item.Height);
            }

            long clampedDx =
                Math.Max(
                    -minX,
                    Math.Min(
                        Document.WidthEmu -
                        maxRight,
                        dx));
            long clampedDy =
                Math.Max(
                    -minY,
                    Math.Min(
                        Document.HeightEmu -
                        maxBottom,
                        dy));

            for (int i = 0;
                 i < snapshot.Count;
                 i++)
            {
                SelectionGeometrySnapshot item =
                    snapshot[i];

                SetObjectGeometry(
                    item.Kind,
                    item.Index,
                    item.X +
                    clampedDx,
                    item.Y +
                    clampedDy,
                    item.Width,
                    item.Height);
            }
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

            slide.SynchronizeObjectOrder();

            for (int orderIndex =
                     slide.ObjectOrder.Count - 1;
                 orderIndex >= 0;
                 orderIndex--)
            {
                PresentationLayerEntry entry =
                    slide.ObjectOrder[orderIndex];

                if (entry == null)
                    continue;

                RectangleF candidate =
                    RectangleF.Empty;
                EditorObjectKind candidateKind =
                    EditorObjectKind.None;

                if (entry.Kind ==
                        PresentationLayerKind.TextBox &&
                    entry.Index >= 0 &&
                    entry.Index < slide.TextBoxes.Count)
                {
                    PresentationTextBox item =
                        slide.TextBoxes[entry.Index];
                    candidate = ToRectangle(
                        item.X,
                        item.Y,
                        item.Width,
                        item.Height);
                    candidateKind =
                        EditorObjectKind.TextBox;
                }
                else if (entry.Kind ==
                             PresentationLayerKind.Image &&
                         entry.Index >= 0 &&
                         entry.Index < slide.Images.Count)
                {
                    PresentationImage item =
                        slide.Images[entry.Index];
                    candidate = ToRectangle(
                        item.X,
                        item.Y,
                        item.Width,
                        item.Height);
                    candidateKind =
                        EditorObjectKind.Image;
                }
                else if (entry.Kind ==
                             PresentationLayerKind.Shape &&
                         entry.Index >= 0 &&
                         entry.Index < slide.Shapes.Count)
                {
                    PresentationShape item =
                        slide.Shapes[entry.Index];
                    candidate = ToRectangle(
                        item.X,
                        item.Y,
                        item.Width,
                        item.Height);
                    candidateKind =
                        EditorObjectKind.Shape;
                }

                if (candidateKind !=
                        EditorObjectKind.None &&
                    candidate.Contains(point))
                {
                    kind = candidateKind;
                    index = entry.Index;
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
            return GetObjectGeometry(
                SelectedObjectKind,
                SelectedObjectIndex,
                out x,
                out y,
                out width,
                out height);
        }

        private bool GetObjectGeometry(
            EditorObjectKind kind,
            int index,
            out long x,
            out long y,
            out long width,
            out long height)
        {
            x = 0;
            y = 0;
            width = 1;
            height = 1;

            PresentationSlide slide =
                GetSlide();

            if (slide == null ||
                index < 0)
            {
                return false;
            }

            if (kind == EditorObjectKind.TextBox &&
                index < slide.TextBoxes.Count)
            {
                PresentationTextBox item =
                    slide.TextBoxes[index];
                x = item.X;
                y = item.Y;
                width = item.Width;
                height = item.Height;
                return true;
            }

            if (kind == EditorObjectKind.Shape &&
                index < slide.Shapes.Count)
            {
                PresentationShape item =
                    slide.Shapes[index];
                x = item.X;
                y = item.Y;
                width = item.Width;
                height = item.Height;
                return true;
            }

            if (kind == EditorObjectKind.Image &&
                index < slide.Images.Count)
            {
                PresentationImage item =
                    slide.Images[index];
                x = item.X;
                y = item.Y;
                width = item.Width;
                height = item.Height;
                return true;
            }

            return false;
        }

        private void SetSelectedGeometry(
            long x,
            long y,
            long width,
            long height)
        {
            SetObjectGeometry(
                SelectedObjectKind,
                SelectedObjectIndex,
                x,
                y,
                width,
                height);
        }

        private void SetObjectGeometry(
            EditorObjectKind kind,
            int index,
            long x,
            long y,
            long width,
            long height)
        {
            PresentationSlide slide =
                GetSlide();

            if (slide == null ||
                index < 0)
            {
                return;
            }

            if (kind == EditorObjectKind.TextBox &&
                index < slide.TextBoxes.Count)
            {
                PresentationTextBox item =
                    slide.TextBoxes[index];
                item.X = x;
                item.Y = y;
                item.Width = width;
                item.Height = height;
                return;
            }

            if (kind == EditorObjectKind.Shape &&
                index < slide.Shapes.Count)
            {
                PresentationShape item =
                    slide.Shapes[index];
                item.X = x;
                item.Y = y;
                item.Width = width;
                item.Height = height;
                return;
            }

            if (kind == EditorObjectKind.Image &&
                index < slide.Images.Count)
            {
                PresentationImage item =
                    slide.Images[index];
                item.X = x;
                item.Y = y;
                item.Width = width;
                item.Height = height;
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
