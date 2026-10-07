using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace PptxViewer
{
    internal static class AdvancedEditorTableExtension
    {
        public static void Attach(
            AdvancedPresentationEditorForm editor,
            PresentationEditSession session)
        {
            if (editor == null || session == null)
                return;

            ListBox slideList = FindControl<ListBox>(editor);
            AdvancedPresentationCanvas canvas =
                FindControl<AdvancedPresentationCanvas>(editor);

            if (slideList == null || canvas == null)
                return;

            EventHandler refresh = delegate
            {
                RefreshTablePreviews(
                    editor,
                    session,
                    slideList,
                    canvas);
            };

            canvas.MouseDoubleClick +=
                delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Left ||
                        canvas.SelectedObjectKind !=
                            EditorObjectKind.Table ||
                        canvas.SelectedObjectIndex < 0)
                    {
                        return;
                    }

                    using (PresentationTableManagerForm form =
                        new PresentationTableManagerForm(
                            session,
                            slideList.SelectedIndex,
                            canvas.SelectedObjectIndex))
                    {
                        form.ShowDialog(editor);
                    }

                    RefreshTablePreviews(
                        editor,
                        session,
                        slideList,
                        canvas);
                };

            slideList.SelectedIndexChanged += refresh;
            canvas.Resize += refresh;
            editor.Shown += refresh;
            refresh(editor, EventArgs.Empty);
        }

        private static void RefreshTablePreviews(
            AdvancedPresentationEditorForm editor,
            PresentationEditSession session,
            ListBox slideList,
            AdvancedPresentationCanvas canvas)
        {
            // Older builds used child controls for table previews. Tables now
            // participate in the canvas' normal paint/hit-test/layer pipeline,
            // so remove any stale preview controls and let the canvas render.
            for (int i = canvas.Controls.Count - 1;
                 i >= 0;
                 i--)
            {
                TablePreviewControl existing =
                    canvas.Controls[i] as TablePreviewControl;

                if (existing != null)
                {
                    canvas.Controls.RemoveAt(i);
                    existing.Dispose();
                }
            }

            canvas.Invalidate();
        }

        private static Rectangle CalculateSlideRectangle(
            Control canvas,
            PresentationDocument document)
        {
            int availableWidth = Math.Max(100, canvas.ClientSize.Width - 48);
            int availableHeight = Math.Max(100, canvas.ClientSize.Height - 48);
            long widthEmu = document.WidthEmu > 0
                ? document.WidthEmu
                : PresentationDocument.DefaultWidthEmu;
            long heightEmu = document.HeightEmu > 0
                ? document.HeightEmu
                : PresentationDocument.DefaultHeightEmu;
            float ratio = widthEmu / (float)Math.Max(1L, heightEmu);
            int width = availableWidth;
            int height = (int)Math.Round(width / ratio);

            if (height > availableHeight)
            {
                height = availableHeight;
                width = (int)Math.Round(height * ratio);
            }

            return new Rectangle(
                (canvas.ClientSize.Width - width) / 2,
                (canvas.ClientSize.Height - height) / 2,
                Math.Max(1, width),
                Math.Max(1, height));
        }

        private static Rectangle ToPixels(
            long x,
            long y,
            long width,
            long height,
            Rectangle slideBounds,
            PresentationDocument document)
        {
            long widthEmu = document.WidthEmu > 0
                ? document.WidthEmu
                : PresentationDocument.DefaultWidthEmu;
            long heightEmu = document.HeightEmu > 0
                ? document.HeightEmu
                : PresentationDocument.DefaultHeightEmu;

            int left = slideBounds.Left +
                (int)Math.Round((x / (double)widthEmu) * slideBounds.Width);
            int top = slideBounds.Top +
                (int)Math.Round((y / (double)heightEmu) * slideBounds.Height);
            int w = (int)Math.Round((width / (double)widthEmu) * slideBounds.Width);
            int h = (int)Math.Round((height / (double)heightEmu) * slideBounds.Height);

            return new Rectangle(left, top, Math.Max(24, w), Math.Max(20, h));
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

        private static Panel FindToolbar(Control root)
        {
            if (root == null)
                return null;

            Panel panel = root as Panel;
            if (panel != null)
            {
                bool hasImage = false;
                bool hasShape = false;

                for (int i = 0; i < panel.Controls.Count; i++)
                {
                    Button button = panel.Controls[i] as Button;
                    if (button == null)
                        continue;

                    if (button.Text == "Image")
                        hasImage = true;
                    else if (button.Text == "Shape")
                        hasShape = true;
                }

                if (hasImage && hasShape)
                    return panel;
            }

            for (int i = 0; i < root.Controls.Count; i++)
            {
                Panel found = FindToolbar(root.Controls[i]);
                if (found != null)
                    return found;
            }

            return null;
        }
    }

    internal sealed class TablePreviewControl : Control
    {
        public PresentationTable Table { get; set; }

        public TablePreviewControl()
        {
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.UserPaint,
                true);
            BackColor = Color.White;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            PresentationTable table = Table;
            if (table == null)
                return;

            int rows = Math.Max(1, table.Rows);
            int columns = Math.Max(1, table.Columns);
            float cellWidth = ClientSize.Width / (float)columns;
            float cellHeight = ClientSize.Height / (float)rows;

            using (Pen border = new Pen(Color.FromArgb(166, 172, 184)))
            {
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        PresentationTableCell cell = table.GetCell(row, column);
                        RectangleF rect = new RectangleF(
                            column * cellWidth,
                            row * cellHeight,
                            cellWidth,
                            cellHeight);

                        Color fill = ParseColor(
                            cell == null ? null : cell.FillColorHex,
                            Color.White);
                        Color text = ParseColor(
                            cell == null ? null : cell.TextColorHex,
                            Color.FromArgb(32, 36, 42));

                        using (Brush fillBrush = new SolidBrush(fill))
                            e.Graphics.FillRectangle(fillBrush, rect);

                        e.Graphics.DrawRectangle(
                            border,
                            Rectangle.Round(rect));

                        if (cell != null && !string.IsNullOrEmpty(cell.Text))
                        {
                            FontStyle style = cell.Bold
                                ? FontStyle.Bold
                                : FontStyle.Regular;
                            float displaySize = Math.Max(
                                6f,
                                Math.Min(18f, cell.FontSizePoints * 0.72f));

                            using (Font font = SafeFont(
                                cell.FontFamily,
                                displaySize,
                                style))
                            using (Brush textBrush = new SolidBrush(text))
                            using (StringFormat format = new StringFormat())
                            {
                                format.Alignment =
                                    cell.Alignment == PresentationTextAlignment.Center
                                    ? StringAlignment.Center
                                    : cell.Alignment == PresentationTextAlignment.Right
                                        ? StringAlignment.Far
                                        : StringAlignment.Near;
                                format.LineAlignment = StringAlignment.Center;
                                format.Trimming = StringTrimming.EllipsisCharacter;
                                e.Graphics.DrawString(
                                    cell.Text,
                                    font,
                                    textBrush,
                                    rect,
                                    format);
                            }
                        }
                    }
                }
            }

            using (Pen outline = new Pen(ApplicationTheme.Accent, 1.5f))
                e.Graphics.DrawRectangle(
                    outline,
                    0,
                    0,
                    Math.Max(0, Width - 1),
                    Math.Max(0, Height - 1));
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
    }

    internal sealed class PresentationTableManagerForm : Form
    {
        private readonly PresentationEditSession session;
        private readonly int slideIndex;
        private readonly ListBox tableList;
        private readonly DataGridView grid;
        private readonly Label infoLabel;
        private bool loadingGrid;

        public PresentationTableManagerForm(
            PresentationEditSession editSession,
            int selectedSlideIndex)
            : this(editSession, selectedSlideIndex, -1)
        {
        }

        public PresentationTableManagerForm(
            PresentationEditSession editSession,
            int selectedSlideIndex,
            int initialTableIndex)
        {
            if (editSession == null)
                throw new ArgumentNullException("editSession");

            session = editSession;
            slideIndex = selectedSlideIndex;

            Text = "Table Editor";
            StartPosition = FormStartPosition.CenterParent;
            Width = 900;
            Height = 620;
            MinimumSize = new Size(700, 480);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);

            Panel toolbar = new Panel();
            toolbar.Dock = DockStyle.Top;
            toolbar.Height = 48;
            toolbar.Padding = new Padding(8);
            toolbar.BackColor = ApplicationTheme.Toolbar;
            Controls.Add(toolbar);

            Button add = MakeButton("Add Table", 8, 84);
            Button delete = MakeButton("Delete", 98, 72);
            Button header = MakeButton("Header Style", 176, 96);
            toolbar.Controls.Add(add);
            toolbar.Controls.Add(delete);
            toolbar.Controls.Add(header);

            infoLabel = new Label();
            infoLabel.Dock = DockStyle.Bottom;
            infoLabel.Height = 28;
            infoLabel.Padding = new Padding(10, 0, 10, 0);
            infoLabel.TextAlign = ContentAlignment.MiddleLeft;
            infoLabel.BackColor = ApplicationTheme.Toolbar;
            infoLabel.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(infoLabel);

            Panel left = new Panel();
            left.Dock = DockStyle.Left;
            left.Width = 180;
            left.Padding = new Padding(8);
            left.BackColor = ApplicationTheme.Sidebar;
            Controls.Add(left);

            tableList = new ListBox();
            tableList.Dock = DockStyle.Fill;
            tableList.BorderStyle = BorderStyle.None;
            tableList.BackColor = ApplicationTheme.Sidebar;
            tableList.ForeColor = ApplicationTheme.PrimaryText;
            left.Controls.Add(tableList);

            grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.RowHeadersVisible = true;
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            Controls.Add(grid);
            grid.BringToFront();

            add.Click += delegate { AddTable(); };
            delete.Click += delegate { DeleteTable(); };
            header.Click += delegate { ApplyHeaderStyle(); };
            tableList.SelectedIndexChanged += delegate { LoadGrid(); };
            grid.CellEndEdit += delegate { SaveGridCell(); };

            RefreshList();
            if (tableList.Items.Count > 0)
            {
                tableList.SelectedIndex =
                    initialTableIndex >= 0 && initialTableIndex < tableList.Items.Count
                    ? initialTableIndex
                    : 0;
            }
            UpdateInfo();
        }

        private void AddTable()
        {
            using (TableSizeDialog dialog = new TableSizeDialog())
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                PresentationTable table = session.AddTable(
                    slideIndex,
                    dialog.Rows,
                    dialog.Columns);
                if (table == null)
                    return;

                int count = GetSlide().Tables.Count;
                table.Name = "Table " + count.ToString();
                table.X = 1371600;
                table.Y = 1828800;
                table.Width = 9448800;
                table.Height = Math.Max(
                    1371600L,
                    dialog.Rows * 685800L);

                RefreshList();
                tableList.SelectedIndex = count - 1;
                UpdateInfo();
            }
        }

        private void DeleteTable()
        {
            int index = tableList.SelectedIndex;
            if (index < 0)
                return;

            if (!session.RemoveTable(slideIndex, index))
                return;

            RefreshList();
            if (tableList.Items.Count > 0)
                tableList.SelectedIndex = Math.Min(index, tableList.Items.Count - 1);
            UpdateInfo();
        }

        private void ApplyHeaderStyle()
        {
            PresentationTable table = GetSelectedTable();
            if (table == null)
                return;

            for (int column = 0; column < table.Columns; column++)
            {
                PresentationTableCell cell = table.GetCell(0, column);
                if (cell == null)
                    continue;

                cell.Bold = true;
                cell.FillColorHex = "EAF0FF";
                cell.TextColorHex = "20242A";
                cell.Alignment = PresentationTextAlignment.Center;
            }

            session.MarkDirty();
            grid.Invalidate();
            UpdateInfo();
        }

        private void RefreshList()
        {
            tableList.Items.Clear();
            PresentationSlide slide = GetSlide();
            if (slide == null)
                return;

            for (int i = 0; i < slide.Tables.Count; i++)
            {
                PresentationTable table = slide.Tables[i];
                string name = table == null || string.IsNullOrEmpty(table.Name)
                    ? "Table " + (i + 1).ToString()
                    : table.Name;
                tableList.Items.Add(name);
            }
        }

        private void LoadGrid()
        {
            loadingGrid = true;

            try
            {
                grid.Rows.Clear();
                grid.Columns.Clear();

                PresentationTable table = GetSelectedTable();
                if (table == null)
                    return;

                for (int column = 0; column < table.Columns; column++)
                    grid.Columns.Add(
                        "Column" + column.ToString(),
                        (column + 1).ToString());

                grid.Rows.Add(table.Rows);

                for (int row = 0; row < table.Rows; row++)
                {
                    grid.Rows[row].HeaderCell.Value = (row + 1).ToString();

                    for (int column = 0; column < table.Columns; column++)
                    {
                        PresentationTableCell cell = table.GetCell(row, column);
                        grid.Rows[row].Cells[column].Value =
                            cell == null ? string.Empty : cell.Text;
                    }
                }
            }
            finally
            {
                loadingGrid = false;
            }

            UpdateInfo();
        }

        private void SaveGridCell()
        {
            if (loadingGrid)
                return;

            PresentationTable table = GetSelectedTable();
            if (table == null || grid.CurrentCell == null)
                return;

            int row = grid.CurrentCell.RowIndex;
            int column = grid.CurrentCell.ColumnIndex;
            PresentationTableCell cell = table.GetCell(row, column);
            if (cell == null)
                return;

            object value = grid.CurrentCell.Value;
            cell.Text = value == null ? string.Empty : value.ToString();
            session.MarkDirty();
            UpdateInfo();
        }

        private PresentationSlide GetSlide()
        {
            if (slideIndex < 0 || slideIndex >= session.Document.Slides.Count)
                return null;
            return session.Document.Slides[slideIndex];
        }

        private PresentationTable GetSelectedTable()
        {
            PresentationSlide slide = GetSlide();
            int index = tableList.SelectedIndex;

            if (slide == null || index < 0 || index >= slide.Tables.Count)
                return null;

            return slide.Tables[index];
        }

        private void UpdateInfo()
        {
            PresentationTable table = GetSelectedTable();
            infoLabel.Text = table == null
                ? "No table selected"
                : table.Rows.ToString() + " row(s) × " +
                    table.Columns.ToString() + " column(s)  •  Double-click a table on the slide to reopen this editor.";
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
    }

    internal sealed class TableSizeDialog : Form
    {
        private readonly NumericUpDown rowCount;
        private readonly NumericUpDown columnCount;

        public int Rows
        {
            get { return (int)rowCount.Value; }
        }

        public int Columns
        {
            get { return (int)columnCount.Value; }
        }

        public TableSizeDialog()
        {
            Text = "New Table";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            Width = 340;
            Height = 210;
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);

            Label rowsLabel = new Label();
            rowsLabel.Text = "Rows";
            rowsLabel.Left = 18;
            rowsLabel.Top = 24;
            rowsLabel.Width = 100;
            Controls.Add(rowsLabel);

            rowCount = new NumericUpDown();
            rowCount.Left = 140;
            rowCount.Top = 20;
            rowCount.Width = 150;
            rowCount.Minimum = 1;
            rowCount.Maximum = 20;
            rowCount.Value = 3;
            Controls.Add(rowCount);

            Label columnsLabel = new Label();
            columnsLabel.Text = "Columns";
            columnsLabel.Left = 18;
            columnsLabel.Top = 64;
            columnsLabel.Width = 100;
            Controls.Add(columnsLabel);

            columnCount = new NumericUpDown();
            columnCount.Left = 140;
            columnCount.Top = 60;
            columnCount.Width = 150;
            columnCount.Minimum = 1;
            columnCount.Maximum = 20;
            columnCount.Value = 3;
            Controls.Add(columnCount);

            Button ok = new Button();
            ok.Text = "Create";
            ok.Left = 134;
            ok.Top = 112;
            ok.Width = 76;
            ok.Height = 32;
            ok.DialogResult = DialogResult.OK;
            ApplicationTheme.ApplyButton(ok);
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancel";
            cancel.Left = 216;
            cancel.Top = 112;
            cancel.Width = 76;
            cancel.Height = 32;
            cancel.DialogResult = DialogResult.Cancel;
            ApplicationTheme.ApplyButton(cancel);
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }
    }
}
