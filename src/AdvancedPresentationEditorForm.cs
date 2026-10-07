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
        Image,
        Table
    }

    internal enum EditorAlignmentCommand
    {
        Left,
        Center,
        Right,
        Top,
        Middle,
        Bottom,
        DistributeHorizontal,
        DistributeVertical
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
        private readonly RichTextBox textEditor;
        private readonly ComboBox fontPicker;
        private readonly NumericUpDown fontSize;
        private readonly CheckBox boldCheck;
        private readonly CheckBox italicCheck;
        private readonly CheckBox underlineCheck;
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
        private readonly Button alignButton;

        private bool loadingProperties;
        private bool propertyEditSnapshotActive;
        private EditorClipboardPackage copiedObjects;

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
                            830,
                            1024,
                            1106
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
            Button addTableButton = MakeButton("Table", 534, 68);
            deleteObjectButton = MakeButton("Delete", 614, 70);
            copyButton = MakeButton("Copy", 690, 64);
            pasteButton = MakeButton("Paste", 760, 64);
            sendBackButton = MakeButton("Send Back", 836, 84);
            bringFrontButton = MakeButton("Bring Front", 926, 92);
            alignButton = MakeButton("Align", 1030, 70);

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
            mainToolbar.Controls.Add(addTableButton);
            mainToolbar.Controls.Add(deleteObjectButton);
            mainToolbar.Controls.Add(copyButton);
            mainToolbar.Controls.Add(pasteButton);
            mainToolbar.Controls.Add(sendBackButton);
            mainToolbar.Controls.Add(bringFrontButton);
            mainToolbar.Controls.Add(alignButton);

            ContextMenuStrip alignMenu =
                new ContextMenuStrip();
            ToolStripMenuItem alignLeft =
                new ToolStripMenuItem("Align left");
            ToolStripMenuItem alignCenter =
                new ToolStripMenuItem("Align center");
            ToolStripMenuItem alignRight =
                new ToolStripMenuItem("Align right");
            ToolStripMenuItem alignTop =
                new ToolStripMenuItem("Align top");
            ToolStripMenuItem alignMiddle =
                new ToolStripMenuItem("Align middle");
            ToolStripMenuItem alignBottom =
                new ToolStripMenuItem("Align bottom");
            ToolStripMenuItem distributeHorizontal =
                new ToolStripMenuItem("Distribute horizontally");
            ToolStripMenuItem distributeVertical =
                new ToolStripMenuItem("Distribute vertically");

            alignMenu.Items.Add(alignLeft);
            alignMenu.Items.Add(alignCenter);
            alignMenu.Items.Add(alignRight);
            alignMenu.Items.Add(new ToolStripSeparator());
            alignMenu.Items.Add(alignTop);
            alignMenu.Items.Add(alignMiddle);
            alignMenu.Items.Add(alignBottom);
            alignMenu.Items.Add(new ToolStripSeparator());
            alignMenu.Items.Add(distributeHorizontal);
            alignMenu.Items.Add(distributeVertical);
            ApplicationTheme.ApplyContextMenu(
                alignMenu);

            alignLeft.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.Left);
            };
            alignCenter.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.Center);
            };
            alignRight.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.Right);
            };
            alignTop.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.Top);
            };
            alignMiddle.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.Middle);
            };
            alignBottom.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.Bottom);
            };
            distributeHorizontal.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.DistributeHorizontal);
            };
            distributeVertical.Click += delegate
            {
                AlignSelectedObjects(
                    EditorAlignmentCommand.DistributeVertical);
            };

            alignMenu.Opening += delegate
            {
                bool canDistribute =
                    canvas != null &&
                    canvas.SelectedObjectCount >= 3;
                distributeHorizontal.Enabled =
                    canDistribute;
                distributeVertical.Enabled =
                    canDistribute;
            };

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
                            alignButton.Right +
                            gap,
                            mainToolbar.ClientSize.Width -
                            10);

                    int availableSelectorWidth =
                        selectorRight -
                        alignButton.Right -
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
            addTableButton.Click += delegate
            {
                OpenTableEditor(-1);
            };
            deleteObjectButton.Click += delegate { DeleteSelectedObject(); };
            copyButton.Click += delegate { CopySelectedObject(); };
            pasteButton.Click += delegate { PasteObject(); };
            sendBackButton.Click += delegate { MoveSelectedObjectLayer(false); };
            bringFrontButton.Click += delegate { MoveSelectedObjectLayer(true); };
            alignButton.Click += delegate
            {
                alignMenu.Show(
                    alignButton,
                    new Point(
                        0,
                        alignButton.Height));
            };

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

            textEditor = new RichTextBox();
            textEditor.Left = 0;
            textEditor.Top = 92;
            textEditor.Width = 270;
            textEditor.Height = 108;
            textEditor.Multiline = true;
            textEditor.ScrollBars = RichTextBoxScrollBars.Vertical;
            textEditor.DetectUrls = false;
            textEditor.HideSelection = false;
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
            fontSize.Width = 76;
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
            boldCheck.Left = 86;
            boldCheck.Top = 288;
            boldCheck.Width = 56;
            boldCheck.ThreeState = true;
            boldCheck.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(boldCheck);

            italicCheck = new CheckBox();
            italicCheck.Text = "Italic";
            italicCheck.Left = 144;
            italicCheck.Top = 288;
            italicCheck.Width = 56;
            italicCheck.ThreeState = true;
            italicCheck.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(italicCheck);

            underlineCheck = new CheckBox();
            underlineCheck.Text = "U";
            underlineCheck.Left = 204;
            underlineCheck.Top = 288;
            underlineCheck.Width = 48;
            underlineCheck.ThreeState = true;
            underlineCheck.ForeColor = ApplicationTheme.PrimaryText;
            properties.Controls.Add(underlineCheck);

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
                        86;
                    italicCheck.Left =
                        contentLeft +
                        144;
                    underlineCheck.Left =
                        contentLeft +
                        204;

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

            ContextMenuStrip canvasMenu =
                new ContextMenuStrip();
            ToolStripMenuItem cutMenuItem =
                new ToolStripMenuItem("Cut");
            ToolStripMenuItem copyMenuItem =
                new ToolStripMenuItem("Copy");
            ToolStripMenuItem pasteMenuItem =
                new ToolStripMenuItem("Paste");
            ToolStripMenuItem duplicateMenuItem =
                new ToolStripMenuItem("Duplicate");
            ToolStripMenuItem deleteMenuItem =
                new ToolStripMenuItem("Delete");
            ToolStripMenuItem selectAllMenuItem =
                new ToolStripMenuItem("Select All");

            cutMenuItem.ShortcutKeyDisplayString = "Ctrl+X";
            copyMenuItem.ShortcutKeyDisplayString = "Ctrl+C";
            pasteMenuItem.ShortcutKeyDisplayString = "Ctrl+V";
            duplicateMenuItem.ShortcutKeyDisplayString = "Ctrl+D";
            deleteMenuItem.ShortcutKeyDisplayString = "Del";
            selectAllMenuItem.ShortcutKeyDisplayString = "Ctrl+A";

            canvasMenu.Items.Add(cutMenuItem);
            canvasMenu.Items.Add(copyMenuItem);
            canvasMenu.Items.Add(pasteMenuItem);
            canvasMenu.Items.Add(duplicateMenuItem);
            canvasMenu.Items.Add(new ToolStripSeparator());
            canvasMenu.Items.Add(deleteMenuItem);
            canvasMenu.Items.Add(new ToolStripSeparator());
            canvasMenu.Items.Add(selectAllMenuItem);
            ApplicationTheme.ApplyContextMenu(
                canvasMenu);
            canvas.ContextMenuStrip = canvasMenu;

            cutMenuItem.Click += delegate
            {
                CutSelectedObjects();
            };
            copyMenuItem.Click += delegate
            {
                CopySelectedObject();
            };
            pasteMenuItem.Click += delegate
            {
                PasteObject();
            };
            duplicateMenuItem.Click += delegate
            {
                DuplicateSelectedObjects();
            };
            deleteMenuItem.Click += delegate
            {
                DeleteSelectedObject();
            };
            selectAllMenuItem.Click += delegate
            {
                canvas.SelectAllObjects();
            };

            canvasMenu.Opening += delegate
            {
                bool hasSelection =
                    canvas.SelectedObjectCount > 0;
                PresentationSlide selectedSlide =
                    GetSelectedSlide();
                bool hasObjects =
                    selectedSlide != null &&
                    (selectedSlide.TextBoxes.Count > 0 ||
                     selectedSlide.Shapes.Count > 0 ||
                     selectedSlide.Images.Count > 0 ||
                     selectedSlide.Tables.Count > 0);

                cutMenuItem.Enabled = hasSelection;
                copyMenuItem.Enabled = hasSelection;
                duplicateMenuItem.Enabled = hasSelection;
                deleteMenuItem.Enabled = hasSelection;
                pasteMenuItem.Enabled = CanPasteObject();
                selectAllMenuItem.Enabled = hasObjects;
            };

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

            textEditor.TextChanged += delegate
            {
                ApplyTextContent();
            };
            textEditor.SelectionChanged += delegate
            {
                RefreshTextSelectionControls();
            };
            fontPicker.SelectedIndexChanged += delegate
            {
                ApplyTextFormatting(
                    RichTextFormatField.FontFamily);
            };
            fontSize.ValueChanged += delegate
            {
                ApplyTextFormatting(
                    RichTextFormatField.FontSize);
            };
            boldCheck.CheckStateChanged += delegate
            {
                ApplyTextFormatting(
                    RichTextFormatField.Bold);
            };
            italicCheck.CheckStateChanged += delegate
            {
                ApplyTextFormatting(
                    RichTextFormatField.Italic);
            };
            underlineCheck.CheckStateChanged += delegate
            {
                ApplyTextFormatting(
                    RichTextFormatField.Underline);
            };
            alignmentPicker.SelectedIndexChanged += delegate
            {
                ApplyTextAlignment();
            };
            textColorEditor.TextChanged += delegate
            {
                ApplyTextFormatting(
                    RichTextFormatField.Color);
            };
            shapeKindPicker.SelectedIndexChanged += delegate
            {
                ApplyShapeProperties();
            };
            fillColorEditor.TextChanged += delegate
            {
                ApplyShapeProperties();
            };
            lineColorEditor.TextChanged += delegate
            {
                ApplyShapeProperties();
            };

            textEditor.Leave += delegate { propertyEditSnapshotActive = false; };
            fontPicker.Leave += delegate { propertyEditSnapshotActive = false; };
            fontSize.Leave += delegate { propertyEditSnapshotActive = false; };
            boldCheck.Leave += delegate { propertyEditSnapshotActive = false; };
            italicCheck.Leave += delegate { propertyEditSnapshotActive = false; };
            underlineCheck.Leave += delegate { propertyEditSnapshotActive = false; };
            textColorEditor.Leave += delegate { propertyEditSnapshotActive = false; };
            shapeKindPicker.Leave += delegate { propertyEditSnapshotActive = false; };
            fillColorEditor.Leave += delegate { propertyEditSnapshotActive = false; };
            lineColorEditor.Leave += delegate { propertyEditSnapshotActive = false; };

            FormClosing += OnEditorClosing;
            Activated += delegate
            {
                UpdateButtons();
            };

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
            bool textEditing =
                IsTextEditingControlActive();

            if (keyData == (Keys.Control | Keys.B) &&
                canvas.GetSelectedTextBox() != null)
            {
                ToggleTextBooleanFormatting(
                    RichTextFormatField.Bold);
                return true;
            }

            if (keyData == (Keys.Control | Keys.I) &&
                canvas.GetSelectedTextBox() != null)
            {
                ToggleTextBooleanFormatting(
                    RichTextFormatField.Italic);
                return true;
            }

            if (keyData == (Keys.Control | Keys.U) &&
                canvas.GetSelectedTextBox() != null)
            {
                ToggleTextBooleanFormatting(
                    RichTextFormatField.Underline);
                return true;
            }

            if (textEditing &&
                (keyData == (Keys.Control | Keys.C) ||
                 keyData == (Keys.Control | Keys.X) ||
                 keyData == (Keys.Control | Keys.V) ||
                 keyData == (Keys.Control | Keys.A) ||
                 keyData == (Keys.Control | Keys.D)))
            {
                return base.ProcessCmdKey(
                    ref msg,
                    keyData);
            }

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

            if (keyData == (Keys.Control | Keys.X))
            {
                CutSelectedObjects();
                return true;
            }

            if (keyData == (Keys.Control | Keys.V))
            {
                PasteObject();
                return true;
            }

            if (keyData == (Keys.Control | Keys.D))
            {
                DuplicateSelectedObjects();
                return true;
            }

            if (keyData == (Keys.Control | Keys.A))
            {
                canvas.SelectAllObjects();
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

        private bool IsTextEditingControlActive()
        {
            Control focused =
                FindFocusedControl(this);

            return
                focused is TextBoxBase ||
                focused is ComboBox ||
                focused is NumericUpDown;
        }

        private static Control FindFocusedControl(
            Control root)
        {
            if (root == null ||
                !root.ContainsFocus)
            {
                return null;
            }

            for (int i = 0;
                 i < root.Controls.Count;
                 i++)
            {
                Control child =
                    root.Controls[i];

                if (child != null &&
                    child.ContainsFocus)
                {
                    Control nested =
                        FindFocusedControl(child);
                    return nested ?? child;
                }
            }

            return root.Focused
                ? root
                : null;
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
                    SetImageDefaultSize(
                        item,
                        image.Width,
                        image.Height);
                }
            }
            catch
            {
                SetImageFallbackSize(item);
            }
        }

        private void SetImageDefaultSize(
            PresentationImage item,
            int pixelWidth,
            int pixelHeight)
        {
            if (item == null ||
                pixelWidth <= 0 ||
                pixelHeight <= 0)
            {
                SetImageFallbackSize(item);
                return;
            }

            double ratio =
                pixelWidth /
                (double)Math.Max(
                    1,
                    pixelHeight);
            long maxWidth = 5486400;
            long maxHeight = 3657600;
            long width = maxWidth;
            long height =
                (long)Math.Round(
                    width /
                    ratio);

            if (height > maxHeight)
            {
                height = maxHeight;
                width =
                    (long)Math.Round(
                        height *
                        ratio);
            }

            item.Width =
                Math.Max(
                    914400L,
                    width);
            item.Height =
                Math.Max(
                    914400L,
                    height);
            item.X =
                Math.Max(
                    0L,
                    (session.Document.WidthEmu -
                     item.Width) /
                    2L);
            item.Y =
                Math.Max(
                    0L,
                    (session.Document.HeightEmu -
                     item.Height) /
                    2L);
        }

        private void SetImageFallbackSize(
            PresentationImage item)
        {
            if (item == null)
                return;

            item.X = 1828800;
            item.Y = 1600200;
            item.Width = 5486400;
            item.Height = 3657600;
        }

        private void OpenTableEditor(
            int initialTableIndex)
        {
            int slideIndex =
                canvas.SelectedSlideIndex;

            if (slideIndex < 0 ||
                slideIndex >=
                    session.Document.Slides.Count)
            {
                return;
            }

            using (PresentationTableManagerForm form =
                new PresentationTableManagerForm(
                    session,
                    slideIndex,
                    initialTableIndex))
            {
                form.ShowDialog(this);
            }

            canvas.Document = session.Document;
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
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
            List<int> tableIndexes =
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
                else if (entry.Kind ==
                         EditorObjectKind.Table)
                {
                    tableIndexes.Add(
                        entry.Index);
                }
            }

            textIndexes.Sort();
            shapeIndexes.Sort();
            imageIndexes.Sort();
            tableIndexes.Sort();

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

            for (int i =
                     tableIndexes.Count - 1;
                 i >= 0;
                 i--)
            {
                removed =
                    session.RemoveTable(
                        slideIndex,
                        tableIndexes[i]) ||
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

            if (canvas.SelectedObjectKind ==
                EditorObjectKind.Table)
            {
                kind =
                    PresentationLayerKind.Table;
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

        private void AlignSelectedObjects(
            EditorAlignmentCommand command)
        {
            if (canvas == null ||
                canvas.SelectedObjectCount < 2)
            {
                return;
            }

            if ((command ==
                    EditorAlignmentCommand.DistributeHorizontal ||
                 command ==
                    EditorAlignmentCommand.DistributeVertical) &&
                canvas.SelectedObjectCount < 3)
            {
                return;
            }

            CaptureHistory();

            if (!canvas.AlignSelection(
                    command))
            {
                return;
            }

            propertyEditSnapshotActive = false;
            session.MarkDirty();
            canvas.Invalidate();
            UpdateStatus();
        }

        private void CutSelectedObjects()
        {
            if (canvas == null ||
                canvas.SelectedObjectCount == 0)
            {
                return;
            }

            if (CopySelectedObject())
                DeleteSelectedObject();
        }

        private void DuplicateSelectedObjects()
        {
            PresentationSlide slide =
                GetSelectedSlide();
            List<EditorSelectionEntry> selection =
                canvas == null
                    ? new List<EditorSelectionEntry>()
                    : canvas.GetSelectedObjects();

            if (slide == null ||
                selection.Count == 0)
            {
                return;
            }

            EditorClipboardPackage package =
                EditorClipboardCodec.CreatePackage(
                    slide,
                    selection);

            if (package == null ||
                !package.HasObjects)
            {
                return;
            }

            CaptureHistory();
            PasteClipboardPackage(
                slide,
                package);
            session.MarkDirty();
            canvas.Invalidate();
            LoadSelectedProperties();
            UpdateStatus();
        }

        private bool CopySelectedObject()
        {
            PresentationSlide slide =
                GetSelectedSlide();
            List<EditorSelectionEntry> selection =
                canvas == null
                    ? new List<EditorSelectionEntry>()
                    : canvas.GetSelectedObjects();

            if (slide == null ||
                selection.Count == 0)
            {
                return false;
            }

            EditorClipboardPackage package =
                EditorClipboardCodec.CreatePackage(
                    slide,
                    selection);

            if (package == null ||
                !package.HasObjects)
            {
                return false;
            }

            copiedObjects = package.Clone();
            TryWriteSystemClipboard(package);
            UpdateButtons();
            return true;
        }

        private void PasteObject()
        {
            PresentationSlide slide =
                GetSelectedSlide();

            if (slide == null)
                return;

            EditorClipboardPackage package;
            string externalText;
            byte[] externalImage;
            int externalImageWidth;
            int externalImageHeight;
            bool clipboardAvailable =
                TryReadSystemClipboard(
                    out package,
                    out externalText,
                    out externalImage,
                    out externalImageWidth,
                    out externalImageHeight);

            if (!clipboardAvailable &&
                copiedObjects != null &&
                copiedObjects.HasObjects)
            {
                package =
                    copiedObjects.Clone();
            }

            if (package != null &&
                package.HasObjects)
            {
                CaptureHistory();
                PasteClipboardPackage(
                    slide,
                    package);
                session.MarkDirty();
                canvas.Invalidate();
                LoadSelectedProperties();
                UpdateStatus();
                return;
            }

            if (externalImage != null &&
                externalImage.Length > 0)
            {
                CaptureHistory();
                PresentationImage item =
                    session.AddImage(
                        canvas.SelectedSlideIndex,
                        externalImage,
                        "png",
                        "image/png");

                if (item == null)
                    return;

                item.Name = "Clipboard image";
                SetImageDefaultSize(
                    item,
                    externalImageWidth,
                    externalImageHeight);

                int imageIndex =
                    slide.Images.Count - 1;
                canvas.Document = session.Document;
                canvas.SelectObject(
                    EditorObjectKind.Image,
                    imageIndex);
                canvas.Invalidate();
                LoadSelectedProperties();
                UpdateStatus();
                return;
            }

            if (!string.IsNullOrEmpty(externalText))
            {
                CaptureHistory();
                PresentationTextBox box =
                    session.AddTextBox(
                        canvas.SelectedSlideIndex,
                        externalText);

                if (box == null)
                    return;

                box.Name = "Clipboard text";
                int textIndex =
                    slide.TextBoxes.Count - 1;
                canvas.Document = session.Document;
                canvas.SelectObject(
                    EditorObjectKind.TextBox,
                    textIndex);
                canvas.Invalidate();
                LoadSelectedProperties();
                UpdateStatus();
            }
        }

        private void PasteClipboardPackage(
            PresentationSlide slide,
            EditorClipboardPackage source)
        {
            if (slide == null ||
                source == null ||
                !source.HasObjects)
            {
                return;
            }

            EditorClipboardPackage package =
                source.Clone();
            const long preferredOffset = 228600L;
            long minX = long.MaxValue;
            long minY = long.MaxValue;
            long maxRight = long.MinValue;
            long maxBottom = long.MinValue;

            for (int i = 0;
                 i < package.Objects.Count;
                 i++)
            {
                EditorClipboardObject item =
                    package.Objects[i];
                long x;
                long y;
                long width;
                long height;

                if (item == null ||
                    !item.TryGetGeometry(
                        out x,
                        out y,
                        out width,
                        out height))
                {
                    continue;
                }

                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                maxRight =
                    Math.Max(
                        maxRight,
                        x + width);
                maxBottom =
                    Math.Max(
                        maxBottom,
                        y + height);
            }

            if (minX == long.MaxValue)
                return;

            long groupWidth =
                Math.Max(
                    0L,
                    maxRight - minX);
            long groupHeight =
                Math.Max(
                    0L,
                    maxBottom - minY);
            long targetLeft =
                groupWidth <= session.Document.WidthEmu
                    ? Math.Max(
                        0L,
                        Math.Min(
                            session.Document.WidthEmu -
                            groupWidth,
                            minX + preferredOffset))
                    : 0L;
            long targetTop =
                groupHeight <= session.Document.HeightEmu
                    ? Math.Max(
                        0L,
                        Math.Min(
                            session.Document.HeightEmu -
                            groupHeight,
                            minY + preferredOffset))
                    : 0L;
            long dx = targetLeft - minX;
            long dy = targetTop - minY;

            List<EditorSelectionEntry> pasted =
                new List<EditorSelectionEntry>();

            for (int i = 0;
                 i < package.Objects.Count;
                 i++)
            {
                EditorClipboardObject item =
                    package.Objects[i];

                if (item == null)
                    continue;

                item.Offset(dx, dy);
                EditorSelectionEntry selection =
                    new EditorSelectionEntry();

                if (item.Kind == EditorObjectKind.TextBox &&
                    item.TextBox != null)
                {
                    slide.TextBoxes.Add(
                        item.TextBox.Clone());
                    selection.Kind =
                        EditorObjectKind.TextBox;
                    selection.Index =
                        slide.TextBoxes.Count - 1;
                    slide.RegisterObjectOrder(
                        PresentationLayerKind.TextBox,
                        selection.Index);
                }
                else if (item.Kind == EditorObjectKind.Shape &&
                         item.Shape != null)
                {
                    slide.Shapes.Add(
                        item.Shape.Clone());
                    selection.Kind =
                        EditorObjectKind.Shape;
                    selection.Index =
                        slide.Shapes.Count - 1;
                    slide.RegisterObjectOrder(
                        PresentationLayerKind.Shape,
                        selection.Index);
                }
                else if (item.Kind == EditorObjectKind.Image &&
                         item.Image != null)
                {
                    slide.Images.Add(
                        item.Image.Clone());
                    selection.Kind =
                        EditorObjectKind.Image;
                    selection.Index =
                        slide.Images.Count - 1;
                    slide.RegisterObjectOrder(
                        PresentationLayerKind.Image,
                        selection.Index);
                }
                else if (item.Kind == EditorObjectKind.Table &&
                         item.Table != null)
                {
                    slide.Tables.Add(
                        item.Table.Clone());
                    selection.Kind =
                        EditorObjectKind.Table;
                    selection.Index =
                        slide.Tables.Count - 1;
                    slide.RegisterObjectOrder(
                        PresentationLayerKind.Table,
                        selection.Index);
                }
                else
                {
                    continue;
                }

                pasted.Add(selection);
            }

            canvas.Document = session.Document;
            canvas.SelectObjects(pasted);
        }

        private void TryWriteSystemClipboard(
            EditorClipboardPackage package)
        {
            if (package == null ||
                !package.HasObjects)
            {
                return;
            }

            Bitmap clipboardBitmap = null;

            try
            {
                DataObject data =
                    new DataObject();
                string encoded =
                    EditorClipboardCodec.Encode(
                        package);

                if (!string.IsNullOrEmpty(encoded))
                {
                    data.SetData(
                        EditorClipboardCodec.ClipboardFormat,
                        false,
                        encoded);
                }

                if (package.Objects.Count == 1)
                {
                    EditorClipboardObject item =
                        package.Objects[0];

                    if (item != null &&
                        item.Kind == EditorObjectKind.TextBox &&
                        item.TextBox != null)
                    {
                        data.SetText(
                            item.TextBox.Text ??
                            string.Empty,
                            TextDataFormat.UnicodeText);
                    }
                    else if (item != null &&
                             item.Kind == EditorObjectKind.Image &&
                             item.Image != null &&
                             item.Image.Data != null &&
                             item.Image.Data.Length > 0)
                    {
                        try
                        {
                            using (MemoryStream stream =
                                new MemoryStream(
                                    item.Image.Data,
                                    false))
                            using (Image source =
                                Image.FromStream(stream))
                            {
                                clipboardBitmap =
                                    new Bitmap(
                                        Math.Max(1, source.Width),
                                        Math.Max(1, source.Height));

                                using (Graphics graphics =
                                    Graphics.FromImage(
                                        clipboardBitmap))
                                {
                                    graphics.DrawImageUnscaled(
                                        source,
                                        0,
                                        0);
                                }

                                data.SetImage(
                                    clipboardBitmap);
                            }
                        }
                        catch
                        {
                            if (clipboardBitmap != null)
                            {
                                clipboardBitmap.Dispose();
                                clipboardBitmap = null;
                            }
                        }
                    }
                }

                Clipboard.SetDataObject(
                    data,
                    true);
            }
            catch
            {
            }
            finally
            {
                if (clipboardBitmap != null)
                    clipboardBitmap.Dispose();
            }
        }

        private bool TryReadSystemClipboard(
            out EditorClipboardPackage package,
            out string text,
            out byte[] imageData,
            out int imageWidth,
            out int imageHeight)
        {
            package = null;
            text = null;
            imageData = null;
            imageWidth = 0;
            imageHeight = 0;

            try
            {
                IDataObject data =
                    Clipboard.GetDataObject();

                if (data == null)
                    return true;

                if (data.GetDataPresent(
                        EditorClipboardCodec.ClipboardFormat))
                {
                    object raw =
                        data.GetData(
                            EditorClipboardCodec.ClipboardFormat);
                    string encoded =
                        raw as string;
                    EditorClipboardPackage decoded;

                    if (EditorClipboardCodec.TryDecode(
                            encoded,
                            out decoded))
                    {
                        package = decoded;
                        return true;
                    }
                }

                if (data.GetDataPresent(
                        DataFormats.Bitmap,
                        true))
                {
                    Image image =
                        data.GetData(
                            DataFormats.Bitmap,
                            true) as Image;

                    if (image != null)
                    {
                        imageWidth = image.Width;
                        imageHeight = image.Height;

                        using (MemoryStream stream =
                            new MemoryStream())
                        {
                            image.Save(
                                stream,
                                System.Drawing.Imaging.ImageFormat.Png);
                            imageData =
                                stream.ToArray();
                        }

                        return true;
                    }
                }

                if (data.GetDataPresent(
                        DataFormats.UnicodeText,
                        true))
                {
                    text =
                        data.GetData(
                            DataFormats.UnicodeText,
                            true) as string;
                    return true;
                }

                if (data.GetDataPresent(
                        DataFormats.Text,
                        true))
                {
                    text =
                        data.GetData(
                            DataFormats.Text,
                            true) as string;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool CanPasteObject()
        {
            try
            {
                IDataObject data =
                    Clipboard.GetDataObject();

                if (data != null)
                {
                    return
                        data.GetDataPresent(
                            EditorClipboardCodec.ClipboardFormat) ||
                        data.GetDataPresent(
                            DataFormats.Bitmap,
                            true) ||
                        data.GetDataPresent(
                            DataFormats.UnicodeText,
                            true) ||
                        data.GetDataPresent(
                            DataFormats.Text,
                            true);
                }

                return false;
            }
            catch
            {
                return
                    copiedObjects != null &&
                    copiedObjects.HasObjects;
            }
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
                PresentationTable table = canvas.GetSelectedTable();

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
                                    : table != null
                                        ? "Table"
                                        : "No object selected";

                textEditor.Enabled = textEnabled;
                fontPicker.Enabled = textEnabled;
                fontSize.Enabled = textEnabled;
                boldCheck.Enabled = textEnabled;
                italicCheck.Enabled = textEnabled;
                underlineCheck.Enabled = textEnabled;
                alignmentPicker.Enabled = textEnabled;
                textColorEditor.Enabled = textEnabled;

                shapeKindPicker.Enabled = shapeEnabled;
                fillColorEditor.Enabled = shapeEnabled;
                lineColorEditor.Enabled = shapeEnabled;

                if (text != null)
                {
                    LoadTextEditorFromModel(text);
                    alignmentPicker.SelectedIndex =
                        (int)text.Alignment;
                    RefreshTextSelectionControls();
                }
                else
                {
                    textEditor.Text = string.Empty;
                    fontPicker.SelectedIndex = -1;
                    boldCheck.CheckState = CheckState.Unchecked;
                    italicCheck.CheckState = CheckState.Unchecked;
                    underlineCheck.CheckState = CheckState.Unchecked;
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

        private void EnsurePropertyHistorySnapshot()
        {
            if (propertyEditSnapshotActive)
                return;

            history.Capture(
                session.Document);
            propertyEditSnapshotActive = true;
        }

        private void ApplyTextContent()
        {
            if (loadingProperties)
                return;

            PresentationTextBox text =
                canvas.GetSelectedTextBox();

            if (text == null)
                return;

            string modelText =
                RichTextSelectionEditor.ToModelText(
                    textEditor.Text);

            if (string.Equals(
                    text.Text,
                    modelText,
                    StringComparison.Ordinal))
            {
                return;
            }

            EnsurePropertyHistorySnapshot();

            // Editing the actual string intentionally drops stale run
            // boundaries. Formatting-only changes use the selection-aware
            // path below and therefore preserve rich runs.
            text.Text = modelText;
            session.MarkDirty();

            int selectionStart =
                textEditor.SelectionStart;
            int selectionLength =
                textEditor.SelectionLength;

            ApplyTextEditorVisualFormatting(
                text,
                selectionStart,
                selectionLength);

            canvas.Invalidate();
            UpdateStatus();
        }

        private void ToggleTextBooleanFormatting(
            RichTextFormatField field)
        {
            CheckBox target = null;

            if (field ==
                RichTextFormatField.Bold)
            {
                target = boldCheck;
            }
            else if (field ==
                     RichTextFormatField.Italic)
            {
                target = italicCheck;
            }
            else if (field ==
                     RichTextFormatField.Underline)
            {
                target = underlineCheck;
            }

            if (target == null)
                return;

            bool turnOn =
                target.CheckState ==
                    CheckState.Indeterminate ||
                !target.Checked;

            target.CheckState =
                turnOn
                    ? CheckState.Checked
                    : CheckState.Unchecked;
        }

        private void ApplyTextFormatting(
            RichTextFormatField field)
        {
            if (loadingProperties)
                return;

            PresentationTextBox text =
                canvas.GetSelectedTextBox();

            if (text == null)
                return;

            RichTextFormatChange change =
                new RichTextFormatChange();
            change.Field = field;

            if (field ==
                RichTextFormatField.FontFamily)
            {
                if (fontPicker.SelectedItem == null)
                    return;

                change.FontFamily =
                    fontPicker.SelectedItem.ToString();
            }
            else if (field ==
                     RichTextFormatField.FontSize)
            {
                change.FontSizePoints =
                    (float)fontSize.Value;
            }
            else if (field ==
                     RichTextFormatField.Bold)
            {
                if (boldCheck.CheckState ==
                    CheckState.Indeterminate)
                {
                    return;
                }

                change.BooleanValue =
                    boldCheck.Checked;
            }
            else if (field ==
                     RichTextFormatField.Italic)
            {
                if (italicCheck.CheckState ==
                    CheckState.Indeterminate)
                {
                    return;
                }

                change.BooleanValue =
                    italicCheck.Checked;
            }
            else if (field ==
                     RichTextFormatField.Underline)
            {
                if (underlineCheck.CheckState ==
                    CheckState.Indeterminate)
                {
                    return;
                }

                change.BooleanValue =
                    underlineCheck.Checked;
            }
            else if (field ==
                     RichTextFormatField.Color)
            {
                string color =
                    NormalizeHex(
                        textColorEditor.Text);

                if (color.Length != 6)
                    return;

                change.ColorHex = color;
            }

            EnsurePropertyHistorySnapshot();

            int start =
                textEditor.SelectionStart;
            int length =
                textEditor.SelectionLength;
            string editorText =
                RichTextSelectionEditor.GetEditorText(
                    text);
            bool changed = false;

            if (length > 0)
            {
                changed =
                    RichTextSelectionEditor.ApplySelection(
                        text,
                        start,
                        length,
                        change);

                if (start == 0 &&
                    length >= editorText.Length)
                {
                    RichTextSelectionEditor
                        .ApplyWholeBoxDefaults(
                            text,
                            change);
                }
            }
            else
            {
                RichTextSelectionEditor
                    .ApplyWholeBoxDefaults(
                        text,
                        change);

                if (text.HasRichText &&
                    editorText.Length > 0)
                {
                    changed =
                        RichTextSelectionEditor.ApplySelection(
                            text,
                            0,
                            editorText.Length,
                            change) ||
                        changed;
                }
                else if (field ==
                             RichTextFormatField.Underline &&
                         editorText.Length > 0)
                {
                    changed =
                        RichTextSelectionEditor.ApplySelection(
                            text,
                            0,
                            editorText.Length,
                            change) ||
                        changed;
                }
                else
                {
                    changed = true;
                }
            }

            if (!changed)
                return;

            session.MarkDirty();

            ApplyTextEditorVisualFormatting(
                text,
                start,
                length);
            RefreshTextSelectionControls();
            canvas.Invalidate();
            UpdateStatus();
        }

        private void ApplyTextAlignment()
        {
            if (loadingProperties)
                return;

            PresentationTextBox text =
                canvas.GetSelectedTextBox();

            if (text == null ||
                alignmentPicker.SelectedIndex < 0 ||
                alignmentPicker.SelectedIndex > 2)
            {
                return;
            }

            PresentationTextAlignment alignment =
                (PresentationTextAlignment)
                alignmentPicker.SelectedIndex;

            EnsurePropertyHistorySnapshot();
            RichTextSelectionEditor
                .ApplyParagraphAlignment(
                    text,
                    alignment);
            session.MarkDirty();
            canvas.Invalidate();
            UpdateStatus();
        }

        private void ApplyShapeProperties()
        {
            if (loadingProperties)
                return;

            PresentationShape shape =
                canvas.GetSelectedShape();

            if (shape == null)
                return;

            EnsurePropertyHistorySnapshot();

            if (shapeKindPicker.SelectedIndex >= 0 &&
                shapeKindPicker.SelectedIndex <= 4)
            {
                shape.Kind =
                    (PresentationShapeKind)
                    shapeKindPicker.SelectedIndex;
            }

            string fill =
                NormalizeHex(
                    fillColorEditor.Text);
            string line =
                NormalizeHex(
                    lineColorEditor.Text);

            if (fill.Length == 6)
                shape.FillColorHex = fill;

            if (line.Length == 6)
                shape.LineColorHex = line;

            session.MarkDirty();
            canvas.Invalidate();
            UpdateStatus();
        }

        private void LoadTextEditorFromModel(
            PresentationTextBox text)
        {
            if (text == null)
            {
                textEditor.Text =
                    string.Empty;
                return;
            }

            textEditor.Text =
                RichTextSelectionEditor.GetEditorText(
                    text);
            ApplyTextEditorVisualFormatting(
                text,
                0,
                0);
        }

        private void ApplyTextEditorVisualFormatting(
            PresentationTextBox text,
            int selectionStart,
            int selectionLength)
        {
            if (text == null)
                return;

            bool previousLoading =
                loadingProperties;
            loadingProperties = true;

            try
            {
                int textLength =
                    textEditor.TextLength;
                int safeStart =
                    Math.Max(
                        0,
                        Math.Min(
                            textLength,
                            selectionStart));
                int safeLength =
                    Math.Max(
                        0,
                        Math.Min(
                            textLength -
                            safeStart,
                            selectionLength));

                textEditor.Select(
                    0,
                    textLength);

                using (Font defaultFont =
                    CreateEditorFont(
                        text.FontFamily,
                        text.FontSizePoints,
                        text.Bold,
                        text.Italic,
                        false))
                {
                    textEditor.SelectionFont =
                        defaultFont;
                }

                textEditor.SelectionColor =
                    ParseEditorColor(
                        text.ColorHex,
                        ApplicationTheme.PrimaryText);

                List<RichTextEditorSegment> segments =
                    RichTextSelectionEditor.GetSegments(
                        text);

                for (int i = 0;
                     i < segments.Count;
                     i++)
                {
                    RichTextEditorSegment segment =
                        segments[i];

                    if (segment == null ||
                        segment.Run == null ||
                        segment.Length <= 0 ||
                        segment.Start >=
                            textLength)
                    {
                        continue;
                    }

                    int length =
                        Math.Min(
                            segment.Length,
                            textLength -
                            segment.Start);

                    textEditor.Select(
                        segment.Start,
                        length);

                    using (Font runFont =
                        CreateEditorFont(
                            string.IsNullOrEmpty(
                                segment.Run.FontFamily)
                                ? text.FontFamily
                                : segment.Run.FontFamily,
                            segment.Run.FontSizePoints > 0f
                                ? segment.Run.FontSizePoints
                                : text.FontSizePoints,
                            segment.Run.Bold,
                            segment.Run.Italic,
                            segment.Run.Underline))
                    {
                        textEditor.SelectionFont =
                            runFont;
                    }

                    textEditor.SelectionColor =
                        ParseEditorColor(
                            string.IsNullOrEmpty(
                                segment.Run.ColorHex)
                                ? text.ColorHex
                                : segment.Run.ColorHex,
                            ApplicationTheme.PrimaryText);
                }

                textEditor.Select(
                    safeStart,
                    safeLength);
            }
            finally
            {
                loadingProperties =
                    previousLoading;
            }
        }

        private void RefreshTextSelectionControls()
        {
            if (loadingProperties)
                return;

            PresentationTextBox text =
                canvas.GetSelectedTextBox();

            if (text == null)
                return;

            bool previousLoading =
                loadingProperties;
            loadingProperties = true;

            try
            {
                RichTextSelectionStyle style = null;
                int selectedCharacters =
                    textEditor.SelectionLength;

                if (selectedCharacters > 0)
                {
                    RichTextSelectionEditor
                        .TryGetSelectionStyle(
                            text,
                            textEditor.SelectionStart,
                            selectedCharacters,
                            out style);
                    objectTypeLabel.Text =
                        "Text box  •  " +
                        selectedCharacters.ToString() +
                        (selectedCharacters == 1
                            ? " char selected"
                            : " chars selected");
                }
                else
                {
                    objectTypeLabel.Text =
                        "Text box";

                    string editorText =
                        RichTextSelectionEditor.GetEditorText(
                            text);

                    if (text.HasRichText &&
                        editorText.Length > 0)
                    {
                        RichTextSelectionEditor
                            .TryGetSelectionStyle(
                                text,
                                0,
                                editorText.Length,
                                out style);
                    }
                }

                if (style == null)
                {
                    style =
                        new RichTextSelectionStyle();
                    style.FontFamily =
                        text.FontFamily;
                    style.FontSizePoints =
                        text.FontSizePoints;
                    style.Bold =
                        text.Bold;
                    style.Italic =
                        text.Italic;
                    style.Underline =
                        false;
                    style.ColorHex =
                        text.ColorHex;
                }

                if (style.FontFamilyMixed)
                    fontPicker.SelectedIndex = -1;
                else
                    SelectFont(
                        style.FontFamily);

                fontSize.Value =
                    (decimal)Math.Max(
                        1f,
                        Math.Min(
                            400f,
                            style.FontSizePoints));

                boldCheck.CheckState =
                    style.BoldMixed
                        ? CheckState.Indeterminate
                        : style.Bold
                            ? CheckState.Checked
                            : CheckState.Unchecked;
                italicCheck.CheckState =
                    style.ItalicMixed
                        ? CheckState.Indeterminate
                        : style.Italic
                            ? CheckState.Checked
                            : CheckState.Unchecked;
                underlineCheck.CheckState =
                    style.UnderlineMixed
                        ? CheckState.Indeterminate
                        : style.Underline
                            ? CheckState.Checked
                            : CheckState.Unchecked;

                textColorEditor.Text =
                    style.ColorMixed
                        ? string.Empty
                        : "#" +
                          NormalizeHex(
                              style.ColorHex);
            }
            finally
            {
                loadingProperties =
                    previousLoading;
            }
        }

        private static Font CreateEditorFont(
            string family,
            float size,
            bool bold,
            bool italic,
            bool underline)
        {
            FontStyle style =
                FontStyle.Regular;

            if (bold)
                style |= FontStyle.Bold;
            if (italic)
                style |= FontStyle.Italic;
            if (underline)
                style |= FontStyle.Underline;

            float safeSize =
                Math.Max(
                    1f,
                    Math.Min(
                        400f,
                        size));

            try
            {
                return new Font(
                    string.IsNullOrEmpty(family)
                        ? "Arial"
                        : family,
                    safeSize,
                    style,
                    GraphicsUnit.Point);
            }
            catch
            {
                return new Font(
                    SystemFonts.MessageBoxFont.FontFamily,
                    safeSize,
                    style,
                    GraphicsUnit.Point);
            }
        }

        private static Color ParseEditorColor(
            string value,
            Color fallback)
        {
            string candidate =
                (value ?? string.Empty)
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
            copyButton.Enabled = hasSelection;
            pasteButton.Enabled = CanPasteObject();

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
            alignButton.Enabled =
                canvas != null &&
                canvas.SelectedObjectCount >= 2;
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

        public void SelectObjects(
            IList<EditorSelectionEntry> objects)
        {
            selectedObjects.Clear();

            if (objects != null)
            {
                for (int i = 0;
                     i < objects.Count;
                     i++)
                {
                    EditorSelectionEntry entry =
                        objects[i];

                    if (entry != null)
                    {
                        AddSelectionEntry(
                            entry.Kind,
                            entry.Index);
                    }
                }
            }

            if (selectedObjects.Count > 0)
            {
                EditorSelectionEntry primary =
                    selectedObjects[
                        selectedObjects.Count - 1];
                SetPrimarySelection(
                    primary.Kind,
                    primary.Index);
            }
            else
            {
                SetPrimarySelection(
                    EditorObjectKind.None,
                    -1);
            }

            NotifySelectionChanged();
        }

        public void SelectAllObjects()
        {
            PresentationSlide slide =
                GetSlide();
            selectedObjects.Clear();

            if (slide != null)
            {
                slide.SynchronizeObjectOrder();

                for (int i = 0;
                     i < slide.ObjectOrder.Count;
                     i++)
                {
                    PresentationLayerEntry layer =
                        slide.ObjectOrder[i];

                    if (layer == null)
                        continue;

                    EditorObjectKind kind =
                        EditorObjectKind.None;

                    if (layer.Kind ==
                        PresentationLayerKind.TextBox)
                    {
                        kind =
                            EditorObjectKind.TextBox;
                    }
                    else if (layer.Kind ==
                             PresentationLayerKind.Shape)
                    {
                        kind =
                            EditorObjectKind.Shape;
                    }
                    else if (layer.Kind ==
                             PresentationLayerKind.Image)
                    {
                        kind =
                            EditorObjectKind.Image;
                    }
                    else if (layer.Kind ==
                             PresentationLayerKind.Table)
                    {
                        kind =
                            EditorObjectKind.Table;
                    }

                    if (kind != EditorObjectKind.None)
                    {
                        AddSelectionEntry(
                            kind,
                            layer.Index);
                    }
                }
            }

            if (selectedObjects.Count > 0)
            {
                EditorSelectionEntry primary =
                    selectedObjects[
                        selectedObjects.Count - 1];
                SetPrimarySelection(
                    primary.Kind,
                    primary.Index);
            }
            else
            {
                SetPrimarySelection(
                    EditorObjectKind.None,
                    -1);
            }

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

        public PresentationTable GetSelectedTable()
        {
            PresentationSlide slide = GetSlide();
            if (slide == null ||
                selectedObjects.Count != 1 ||
                SelectedObjectKind != EditorObjectKind.Table ||
                SelectedObjectIndex < 0 ||
                SelectedObjectIndex >= slide.Tables.Count)
            {
                return null;
            }

            return slide.Tables[SelectedObjectIndex];
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
                else if (entry.Kind ==
                             PresentationLayerKind.Table &&
                         entry.Index >= 0 &&
                         entry.Index < slide.Tables.Count)
                {
                    DrawTable(
                        e.Graphics,
                        slide.Tables[entry.Index],
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

            EditorObjectKind kind;
            int index;
            RectangleF rect;

            if (e.Button == MouseButtons.Right)
            {
                if (HitTest(
                        e.Location,
                        out kind,
                        out index,
                        out rect))
                {
                    if (!IsObjectSelected(
                            kind,
                            index))
                    {
                        SelectObject(
                            kind,
                            index);
                    }
                }
                else
                {
                    ClearSelection();
                }

                return;
            }

            if (e.Button != MouseButtons.Left)
                return;

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
            RectangleF rect =
                ToRectangle(
                    box.X,
                    box.Y,
                    box.Width,
                    box.Height);
            float scale =
                slideRectangle.Width /
                960f;

            PresentationModelTextRenderer.DrawTextBox(
                graphics,
                box,
                rect,
                scale,
                Color.FromArgb(
                    32,
                    36,
                    42));

            DrawSelectionIfNeeded(
                graphics,
                EditorObjectKind.TextBox,
                index,
                rect);
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

        private void DrawTable(
            Graphics graphics,
            PresentationTable table,
            int index)
        {
            if (table == null)
                return;

            RectangleF rect =
                ToRectangle(
                    table.X,
                    table.Y,
                    table.Width,
                    table.Height);
            int rows =
                Math.Max(1, table.Rows);
            int columns =
                Math.Max(1, table.Columns);
            float cellWidth =
                rect.Width / columns;
            float cellHeight =
                rect.Height / rows;
            float scale =
                slideRectangle.Width / 960f;

            using (Pen border =
                new Pen(
                    Color.FromArgb(
                        166,
                        172,
                        184),
                    1f))
            {
                for (int row = 0;
                     row < rows;
                     row++)
                {
                    for (int column = 0;
                         column < columns;
                         column++)
                    {
                        PresentationTableCell cell =
                            table.GetCell(
                                row,
                                column);
                        RectangleF cellRect =
                            new RectangleF(
                                rect.Left +
                                column * cellWidth,
                                rect.Top +
                                row * cellHeight,
                                cellWidth,
                                cellHeight);
                        Color fill =
                            ParseColor(
                                cell == null
                                    ? null
                                    : cell.FillColorHex,
                                Color.White);

                        using (Brush fillBrush =
                            new SolidBrush(fill))
                        {
                            graphics.FillRectangle(
                                fillBrush,
                                cellRect);
                        }

                        graphics.DrawRectangle(
                            border,
                            Rectangle.Round(cellRect));

                        if (cell == null ||
                            string.IsNullOrEmpty(
                                cell.Text))
                        {
                            continue;
                        }

                        FontStyle style =
                            cell.Bold
                                ? FontStyle.Bold
                                : FontStyle.Regular;
                        float displaySize =
                            Math.Max(
                                5f,
                                cell.FontSizePoints *
                                scale);

                        using (Font font =
                            SafeFont(
                                cell.FontFamily,
                                displaySize,
                                style))
                        using (Brush textBrush =
                            new SolidBrush(
                                ParseColor(
                                    cell.TextColorHex,
                                    Color.FromArgb(
                                        32,
                                        36,
                                        42))))
                        using (StringFormat format =
                            new StringFormat())
                        {
                            format.Alignment =
                                cell.Alignment ==
                                    PresentationTextAlignment.Center
                                    ? StringAlignment.Center
                                    : cell.Alignment ==
                                        PresentationTextAlignment.Right
                                        ? StringAlignment.Far
                                        : StringAlignment.Near;
                            format.LineAlignment =
                                StringAlignment.Center;
                            format.Trimming =
                                StringTrimming.EllipsisCharacter;

                            RectangleF textRect =
                                RectangleF.Inflate(
                                    cellRect,
                                    -4f,
                                    -2f);
                            graphics.DrawString(
                                cell.Text,
                                font,
                                textBrush,
                                textRect,
                                format);
                        }
                    }
                }
            }

            DrawSelectionIfNeeded(
                graphics,
                EditorObjectKind.Table,
                index,
                rect);
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



        public bool AlignSelection(
            EditorAlignmentCommand command)
        {
            if (selectedObjects.Count < 2 ||
                Document == null)
            {
                return false;
            }

            List<SelectionGeometrySnapshot> items =
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
                items.Add(item);
            }

            if (items.Count < 2)
                return false;

            if ((command ==
                    EditorAlignmentCommand.DistributeHorizontal ||
                 command ==
                    EditorAlignmentCommand.DistributeVertical) &&
                items.Count < 3)
            {
                return false;
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
                 i < items.Count;
                 i++)
            {
                SelectionGeometrySnapshot item =
                    items[i];

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

            if (command ==
                EditorAlignmentCommand.DistributeHorizontal)
            {
                items.Sort(
                    delegate(
                        SelectionGeometrySnapshot left,
                        SelectionGeometrySnapshot right)
                    {
                        long leftCenter =
                            left.X +
                            left.Width /
                            2L;
                        long rightCenter =
                            right.X +
                            right.Width /
                            2L;
                        return leftCenter.CompareTo(
                            rightCenter);
                    });

                long firstCenter =
                    items[0].X +
                    items[0].Width /
                    2L;
                long lastCenter =
                    items[items.Count - 1].X +
                    items[items.Count - 1].Width /
                    2L;
                double step =
                    (lastCenter -
                     firstCenter) /
                    (double)(
                        items.Count - 1);

                for (int i = 1;
                     i < items.Count - 1;
                     i++)
                {
                    SelectionGeometrySnapshot item =
                        items[i];
                    long center =
                        (long)Math.Round(
                            firstCenter +
                            step *
                            i);
                    long x =
                        center -
                        item.Width /
                        2L;

                    x =
                        Math.Max(
                            0L,
                            Math.Min(
                                Math.Max(
                                    0L,
                                    Document.WidthEmu -
                                    item.Width),
                                x));

                    SetObjectGeometry(
                        item.Kind,
                        item.Index,
                        x,
                        item.Y,
                        item.Width,
                        item.Height);
                }

                return true;
            }

            if (command ==
                EditorAlignmentCommand.DistributeVertical)
            {
                items.Sort(
                    delegate(
                        SelectionGeometrySnapshot top,
                        SelectionGeometrySnapshot bottom)
                    {
                        long topCenter =
                            top.Y +
                            top.Height /
                            2L;
                        long bottomCenter =
                            bottom.Y +
                            bottom.Height /
                            2L;
                        return topCenter.CompareTo(
                            bottomCenter);
                    });

                long firstCenter =
                    items[0].Y +
                    items[0].Height /
                    2L;
                long lastCenter =
                    items[items.Count - 1].Y +
                    items[items.Count - 1].Height /
                    2L;
                double step =
                    (lastCenter -
                     firstCenter) /
                    (double)(
                        items.Count - 1);

                for (int i = 1;
                     i < items.Count - 1;
                     i++)
                {
                    SelectionGeometrySnapshot item =
                        items[i];
                    long center =
                        (long)Math.Round(
                            firstCenter +
                            step *
                            i);
                    long y =
                        center -
                        item.Height /
                        2L;

                    y =
                        Math.Max(
                            0L,
                            Math.Min(
                                Math.Max(
                                    0L,
                                    Document.HeightEmu -
                                    item.Height),
                                y));

                    SetObjectGeometry(
                        item.Kind,
                        item.Index,
                        item.X,
                        y,
                        item.Width,
                        item.Height);
                }

                return true;
            }

            long centerX =
                minX +
                (maxRight -
                 minX) /
                2L;
            long centerY =
                minY +
                (maxBottom -
                 minY) /
                2L;

            for (int i = 0;
                 i < items.Count;
                 i++)
            {
                SelectionGeometrySnapshot item =
                    items[i];
                long x =
                    item.X;
                long y =
                    item.Y;

                if (command ==
                    EditorAlignmentCommand.Left)
                {
                    x = minX;
                }
                else if (command ==
                         EditorAlignmentCommand.Center)
                {
                    x =
                        centerX -
                        item.Width /
                        2L;
                }
                else if (command ==
                         EditorAlignmentCommand.Right)
                {
                    x =
                        maxRight -
                        item.Width;
                }
                else if (command ==
                         EditorAlignmentCommand.Top)
                {
                    y = minY;
                }
                else if (command ==
                         EditorAlignmentCommand.Middle)
                {
                    y =
                        centerY -
                        item.Height /
                        2L;
                }
                else if (command ==
                         EditorAlignmentCommand.Bottom)
                {
                    y =
                        maxBottom -
                        item.Height;
                }

                x =
                    Math.Max(
                        0L,
                        Math.Min(
                            Math.Max(
                                0L,
                                Document.WidthEmu -
                                item.Width),
                            x));
                y =
                    Math.Max(
                        0L,
                        Math.Min(
                            Math.Max(
                                0L,
                                Document.HeightEmu -
                                item.Height),
                            y));

                SetObjectGeometry(
                    item.Kind,
                    item.Index,
                    x,
                    y,
                    item.Width,
                    item.Height);
            }

            return true;
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
                else if (entry.Kind ==
                             PresentationLayerKind.Table &&
                         entry.Index >= 0 &&
                         entry.Index < slide.Tables.Count)
                {
                    PresentationTable item =
                        slide.Tables[entry.Index];
                    candidate = ToRectangle(
                        item.X,
                        item.Y,
                        item.Width,
                        item.Height);
                    candidateKind =
                        EditorObjectKind.Table;
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

            if (kind == EditorObjectKind.Table &&
                index < slide.Tables.Count)
            {
                PresentationTable item =
                    slide.Tables[index];
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
                return;
            }

            if (kind == EditorObjectKind.Table &&
                index < slide.Tables.Count)
            {
                PresentationTable item =
                    slide.Tables[index];
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
            return PresentationRenderPrimitives
                .CreateBasicShapePath(
                    kind,
                    rect);
        }

        private static Font SafeFont(string family, float size, FontStyle style)
        {
            return PresentationRenderPrimitives
                .SafeFont(
                    family,
                    size,
                    style);
        }

        private static Color ParseColor(string value, Color fallback)
        {
            return PresentationRenderPrimitives
                .ParseHexColor(
                    value,
                    fallback);
        }
    }
}
