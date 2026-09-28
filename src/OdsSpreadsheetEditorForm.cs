using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class OdsSpreadsheetEditSession
    {
        public SpreadsheetDocument Document { get; private set; }
        public string FilePath { get; private set; }
        public bool IsDirty { get; private set; }

        private OdsSpreadsheetEditSession(SpreadsheetDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            Document = document;
            FilePath = string.Empty;
            IsDirty = true;
        }

        public static OdsSpreadsheetEditSession CreateNew(string title)
        {
            return new OdsSpreadsheetEditSession(
                SpreadsheetDocument.CreateNew(title));
        }

        public static OdsSpreadsheetEditSession Open(string path)
        {
            OdsEditSafetyResult safety = OdsEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
            {
                throw new InvalidOperationException(
                    safety == null || string.IsNullOrEmpty(safety.Warning)
                        ? "This ODS cannot yet be edited without risking unsupported-content loss."
                        : safety.Warning);
            }

            OdsSpreadsheetEditSession session =
                new OdsSpreadsheetEditSession(OdsReader.Read(path));
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
                throw new InvalidOperationException("Save As is required for a new ODS workbook.");

            OdsWriter.Save(Document, FilePath);
            IsDirty = false;
        }

        public void SaveAs(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("A destination path is required.", "path");

            if (!string.Equals(
                    Path.GetExtension(path),
                    ".ods",
                    StringComparison.OrdinalIgnoreCase))
            {
                path = Path.ChangeExtension(path, "ods");
            }

            OdsWriter.Save(Document, path);
            FilePath = Path.GetFullPath(path);
            IsDirty = false;
        }
    }

    internal sealed class OdsSpreadsheetEditorForm : Form
    {
        private readonly OdsSpreadsheetEditSession session;
        private readonly TabControl sheetTabs;
        private readonly DataGridView grid;
        private readonly TextBox inputBar;
        private readonly Label cellLabel;
        private readonly Label statusLabel;
        private int currentSheetIndex;
        private bool loadingGrid;

        public string SavedFilePath
        {
            get { return session.FilePath; }
        }

        public OdsSpreadsheetEditorForm(OdsSpreadsheetEditSession editSession)
        {
            if (editSession == null)
                throw new ArgumentNullException("editSession");

            session = editSession;

            Text = "ODS Spreadsheet Editor";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1320;
            Height = 820;
            MinimumSize = new Size(900, 600);
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
            Button addSheet = MakeButton("+ Sheet", 166, 74);
            Button deleteSheet = MakeButton("Delete", 246, 70);
            Button renameSheet = MakeButton("Rename", 322, 72);
            Button moveLeft = MakeButton("Left", 400, 58);
            Button moveRight = MakeButton("Right", 464, 62);

            toolbar.Controls.Add(save);
            toolbar.Controls.Add(saveAs);
            toolbar.Controls.Add(addSheet);
            toolbar.Controls.Add(deleteSheet);
            toolbar.Controls.Add(renameSheet);
            toolbar.Controls.Add(moveLeft);
            toolbar.Controls.Add(moveRight);

            save.Click += delegate { SaveDocument(false); };
            saveAs.Click += delegate { SaveDocument(true); };
            addSheet.Click += delegate { AddSheet(); };
            deleteSheet.Click += delegate { DeleteSheet(); };
            renameSheet.Click += delegate { RenameSheet(); };
            moveLeft.Click += delegate { MoveSheet(-1); };
            moveRight.Click += delegate { MoveSheet(1); };

            Label formatBadge = new Label();
            formatBadge.Text = "ODF · ODS";
            formatBadge.AutoSize = true;
            formatBadge.Left = 550;
            formatBadge.Top = 16;
            formatBadge.ForeColor = ApplicationTheme.Accent;
            toolbar.Controls.Add(formatBadge);

            Panel inputPanel = new Panel();
            inputPanel.Dock = DockStyle.Top;
            inputPanel.Height = 38;
            inputPanel.Padding = new Padding(8, 5, 8, 5);
            inputPanel.BackColor = ApplicationTheme.Surface;
            Controls.Add(inputPanel);
            inputPanel.BringToFront();

            cellLabel = new Label();
            cellLabel.Left = 8;
            cellLabel.Top = 8;
            cellLabel.Width = 70;
            cellLabel.Height = 22;
            cellLabel.Text = "A1";
            cellLabel.TextAlign = ContentAlignment.MiddleCenter;
            cellLabel.BackColor = ApplicationTheme.Toolbar;
            cellLabel.ForeColor = ApplicationTheme.PrimaryText;
            inputPanel.Controls.Add(cellLabel);

            inputBar = new TextBox();
            inputBar.Left = 86;
            inputBar.Top = 7;
            inputBar.Width = 900;
            inputBar.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            inputBar.BorderStyle = BorderStyle.FixedSingle;
            inputBar.BackColor = Color.White;
            inputBar.ForeColor = Color.FromArgb(28, 31, 36);
            inputPanel.Controls.Add(inputBar);

            statusLabel = new Label();
            statusLabel.Dock = DockStyle.Bottom;
            statusLabel.Height = 28;
            statusLabel.Padding = new Padding(10, 0, 10, 0);
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusLabel.BackColor = ApplicationTheme.Toolbar;
            statusLabel.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(statusLabel);

            sheetTabs = new TabControl();
            sheetTabs.Dock = DockStyle.Bottom;
            sheetTabs.Height = 34;
            Controls.Add(sheetTabs);

            grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.BackgroundColor = Color.FromArgb(248, 249, 251);
            grid.BorderStyle = BorderStyle.None;
            grid.GridColor = Color.FromArgb(218, 221, 226);
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(232, 235, 240);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(40, 44, 51);
            grid.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(232, 235, 240);
            grid.RowHeadersDefaultCellStyle.ForeColor = Color.FromArgb(70, 75, 84);
            grid.DefaultCellStyle.BackColor = Color.White;
            grid.DefaultCellStyle.ForeColor = Color.FromArgb(28, 31, 36);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(211, 225, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(20, 24, 31);
            grid.RowHeadersWidth = 56;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.MultiSelect = true;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            Controls.Add(grid);
            grid.BringToFront();

            sheetTabs.SelectedIndexChanged += delegate
            {
                if (sheetTabs.SelectedIndex >= 0 &&
                    sheetTabs.SelectedIndex < session.Document.Sheets.Count)
                {
                    currentSheetIndex = sheetTabs.SelectedIndex;
                    LoadSheet();
                }
            };

            grid.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (!loadingGrid && e.RowIndex >= 0 && e.ColumnIndex >= 0)
                    CommitGridCell(e.RowIndex, e.ColumnIndex);
            };

            grid.CellEndEdit += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                    CommitGridCell(e.RowIndex, e.ColumnIndex);
            };

            grid.CurrentCellChanged += delegate
            {
                LoadCurrentCellIntoInputBar();
            };

            inputBar.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    CommitInputBar();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };

            FormClosing += OnEditorClosing;

            RefreshSheetTabs();
            ConfigureGrid();
            LoadSheet();
            UpdateStatus();
        }

        public static OdsSpreadsheetEditorForm CreateNew(string title)
        {
            return new OdsSpreadsheetEditorForm(
                OdsSpreadsheetEditSession.CreateNew(title));
        }

        public static OdsSpreadsheetEditorForm Open(string path)
        {
            return new OdsSpreadsheetEditorForm(
                OdsSpreadsheetEditSession.Open(path));
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

            if (keyData == Keys.Delete && grid.Focused)
            {
                ClearSelectedCells();
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

        private void ConfigureGrid()
        {
            loadingGrid = true;
            try
            {
                grid.Columns.Clear();
                grid.Rows.Clear();

                const int columns = 26;
                const int rows = 100;

                for (int column = 1; column <= columns; column++)
                {
                    DataGridViewTextBoxColumn item = new DataGridViewTextBoxColumn();
                    item.HeaderText = ColumnName(column);
                    item.Width = 90;
                    item.SortMode = DataGridViewColumnSortMode.NotSortable;
                    grid.Columns.Add(item);
                }

                grid.Rows.Add(rows);
                for (int row = 0; row < rows; row++)
                    grid.Rows[row].HeaderCell.Value = (row + 1).ToString();
            }
            finally
            {
                loadingGrid = false;
            }
        }

        private void RefreshSheetTabs()
        {
            sheetTabs.TabPages.Clear();
            for (int i = 0; i < session.Document.Sheets.Count; i++)
            {
                SpreadsheetSheet sheet = session.Document.Sheets[i];
                sheetTabs.TabPages.Add(
                    new TabPage(
                        sheet == null || string.IsNullOrEmpty(sheet.Name)
                            ? "Sheet " + (i + 1).ToString()
                            : sheet.Name));
            }

            if (sheetTabs.TabPages.Count > 0)
            {
                currentSheetIndex = Math.Max(
                    0,
                    Math.Min(currentSheetIndex, sheetTabs.TabPages.Count - 1));
                sheetTabs.SelectedIndex = currentSheetIndex;
            }
        }

        private void LoadSheet()
        {
            if (currentSheetIndex < 0 ||
                currentSheetIndex >= session.Document.Sheets.Count)
                return;

            loadingGrid = true;
            try
            {
                for (int row = 0; row < grid.Rows.Count; row++)
                    for (int column = 0; column < grid.Columns.Count; column++)
                        grid[column, row].Value = null;

                SpreadsheetSheet sheet = session.Document.Sheets[currentSheetIndex];
                if (sheet != null)
                {
                    System.Collections.Generic.List<SpreadsheetCell> cells =
                        sheet.GetCellsSorted();
                    for (int i = 0; i < cells.Count; i++)
                    {
                        SpreadsheetCell cell = cells[i];
                        if (cell == null ||
                            cell.Row < 1 || cell.Row > grid.Rows.Count ||
                            cell.Column < 1 || cell.Column > grid.Columns.Count)
                            continue;

                        grid[cell.Column - 1, cell.Row - 1].Value = cell.DisplayText;
                    }
                }
            }
            finally
            {
                loadingGrid = false;
            }

            LoadCurrentCellIntoInputBar();
            UpdateStatus();
        }

        private void CommitGridCell(int rowIndex, int columnIndex)
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null)
                return;

            object value = grid[columnIndex, rowIndex].Value;
            sheet.SetInput(
                rowIndex + 1,
                columnIndex + 1,
                value == null ? string.Empty : value.ToString());
            session.MarkDirty();
            UpdateStatus();
            LoadCurrentCellIntoInputBar();
        }

        private void LoadCurrentCellIntoInputBar()
        {
            if (grid.CurrentCell == null)
            {
                cellLabel.Text = "";
                inputBar.Text = "";
                return;
            }

            int row = grid.CurrentCell.RowIndex + 1;
            int column = grid.CurrentCell.ColumnIndex + 1;
            cellLabel.Text = ColumnName(column) + row.ToString();

            SpreadsheetSheet sheet = GetCurrentSheet();
            SpreadsheetCell cell = sheet == null
                ? null
                : sheet.GetCell(row, column, false);
            inputBar.Text = cell == null ? string.Empty : cell.DisplayText;
        }

        private void CommitInputBar()
        {
            if (grid.CurrentCell == null)
                return;

            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null)
                return;

            int row = grid.CurrentCell.RowIndex + 1;
            int column = grid.CurrentCell.ColumnIndex + 1;
            sheet.SetInput(row, column, inputBar.Text);
            session.MarkDirty();
            LoadSheet();
            grid.CurrentCell = grid[column - 1, row - 1];
        }

        private void ClearSelectedCells()
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null)
                return;

            foreach (DataGridViewCell selected in grid.SelectedCells)
            {
                if (selected.RowIndex < 0 || selected.ColumnIndex < 0)
                    continue;

                sheet.SetInput(
                    selected.RowIndex + 1,
                    selected.ColumnIndex + 1,
                    string.Empty);
                selected.Value = null;
            }

            session.MarkDirty();
            UpdateStatus();
            LoadCurrentCellIntoInputBar();
        }

        private void AddSheet()
        {
            SpreadsheetSheet sheet = session.Document.AddSheet("Sheet");
            session.MarkDirty();
            RefreshSheetTabs();
            currentSheetIndex = session.Document.Sheets.IndexOf(sheet);
            sheetTabs.SelectedIndex = currentSheetIndex;
            UpdateStatus();
        }

        private void DeleteSheet()
        {
            if (!session.Document.RemoveSheet(currentSheetIndex))
            {
                MessageBox.Show(
                    this,
                    "A workbook must keep at least one worksheet.",
                    "Delete sheet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            session.MarkDirty();
            currentSheetIndex = Math.Min(
                currentSheetIndex,
                session.Document.Sheets.Count - 1);
            RefreshSheetTabs();
            LoadSheet();
        }

        private void RenameSheet()
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null)
                return;

            string name = ViewerDialogs.Prompt(
                this,
                "Worksheet name:",
                "Rename sheet",
                sheet.Name);

            if (string.IsNullOrEmpty(name))
                return;

            if (!session.Document.RenameSheet(currentSheetIndex, name))
            {
                MessageBox.Show(
                    this,
                    "The worksheet name is invalid or already in use.",
                    "Rename sheet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            session.MarkDirty();
            RefreshSheetTabs();
            UpdateStatus();
        }

        private void MoveSheet(int direction)
        {
            int target = currentSheetIndex + direction;
            if (!session.Document.MoveSheet(currentSheetIndex, target))
                return;

            session.MarkDirty();
            currentSheetIndex = target;
            RefreshSheetTabs();
            UpdateStatus();
        }

        private SpreadsheetSheet GetCurrentSheet()
        {
            if (currentSheetIndex < 0 ||
                currentSheetIndex >= session.Document.Sheets.Count)
                return null;
            return session.Document.Sheets[currentSheetIndex];
        }

        private void SaveDocument(bool saveAs)
        {
            try
            {
                if (saveAs || string.IsNullOrEmpty(session.FilePath))
                {
                    using (SaveFileDialog dialog = new SaveFileDialog())
                    {
                        dialog.Filter = "OpenDocument Spreadsheet (*.ods)|*.ods";
                        dialog.DefaultExt = "ods";
                        dialog.AddExtension = true;
                        dialog.FileName = SafeFileName(session.Document.Title) + ".ods";

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
                Text = Path.GetFileName(session.FilePath) + " - ODS Editor";
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
                    ? "Unsaved ODS"
                    : Path.GetFileName(session.FilePath)) +
                "  •  " + session.Document.Sheets.Count.ToString() + " sheet(s)";
        }

        private void OnEditorClosing(object sender, FormClosingEventArgs e)
        {
            if (!session.IsDirty)
                return;

            DialogResult result = MessageBox.Show(
                this,
                "Save changes before closing?",
                "Unsaved ODS workbook",
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

        private static string ColumnName(int column)
        {
            string name = string.Empty;
            int value = Math.Max(1, column);
            while (value > 0)
            {
                value--;
                name = (char)('A' + (value % 26)) + name;
                value /= 26;
            }
            return name;
        }

        private static string SafeFileName(string value)
        {
            string text = string.IsNullOrEmpty(value) ? "New Spreadsheet" : value;
            char[] invalid = Path.GetInvalidFileNameChars();
            for (int i = 0; i < invalid.Length; i++)
                text = text.Replace(invalid[i], '_');
            return text;
        }
    }
}
