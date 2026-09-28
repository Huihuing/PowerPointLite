using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class SpreadsheetEditorForm : Form
    {
        private readonly SpreadsheetEditSession session;
        private readonly TabControl sheetTabs;
        private readonly DataGridView grid;
        private readonly TextBox inputBar;
        private readonly Label cellLabel;
        private readonly Label statusLabel;
        private bool loadingGrid;
        private int currentSheetIndex;

        public string SavedFilePath
        {
            get { return session.FilePath; }
        }

        public SpreadsheetEditorForm(SpreadsheetEditSession editSession)
        {
            if (editSession == null)
                throw new ArgumentNullException("editSession");

            session = editSession;
            currentSheetIndex = 0;

            Text = "Spreadsheet Editor";
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
            sheetTabs.Appearance = TabAppearance.Normal;
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
            grid.AllowUserToOrderColumns = false;
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

            grid.CurrentCellChanged += delegate
            {
                LoadCurrentCellIntoInputBar();
            };

            grid.CellEndEdit += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                    CommitGridCell(e.RowIndex, e.ColumnIndex);
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

        public static SpreadsheetEditorForm CreateNew(string title)
        {
            return new SpreadsheetEditorForm(
                SpreadsheetEditSession.CreateNew(title));
        }

        public static SpreadsheetEditorForm Open(string path)
        {
            return new SpreadsheetEditorForm(
                SpreadsheetEditSession.Open(path));
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

        private void ConfigureGrid()
        {
            loadingGrid = true;
            try
            {
                grid.Columns.Clear();
                grid.Rows.Clear();

                int columns = 26;
                int rows = 100;
                SpreadsheetSheet sheet = GetCurrentSheet();
                if (sheet != null)
                {
                    columns = Math.Max(columns, Math.Min(100, sheet.MaxColumn + 5));
                    rows = Math.Max(rows, Math.Min(2000, sheet.MaxRow + 20));
                }

                for (int column = 1; column <= columns; column++)
                {
                    DataGridViewTextBoxColumn item = new DataGridViewTextBoxColumn();
                    item.Name = "C" + column.ToString(CultureInfo.InvariantCulture);
                    item.HeaderText = ColumnName(column);
                    item.Width = 96;
                    item.SortMode = DataGridViewColumnSortMode.NotSortable;
                    grid.Columns.Add(item);
                }

                grid.Rows.Add(rows);
                for (int row = 0; row < grid.Rows.Count; row++)
                    grid.Rows[row].HeaderCell.Value = (row + 1).ToString(CultureInfo.InvariantCulture);
            }
            finally
            {
                loadingGrid = false;
            }
        }

        private void LoadSheet()
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null)
                return;

            int requiredColumns = Math.Max(26, Math.Min(100, sheet.MaxColumn + 5));
            int requiredRows = Math.Max(100, Math.Min(2000, sheet.MaxRow + 20));

            if (grid.Columns.Count < requiredColumns || grid.Rows.Count < requiredRows)
                ConfigureGrid();

            loadingGrid = true;
            try
            {
                for (int r = 0; r < grid.Rows.Count; r++)
                    for (int c = 0; c < grid.Columns.Count; c++)
                        grid.Rows[r].Cells[c].Value = null;

                System.Collections.Generic.List<SpreadsheetCell> cells = sheet.GetCellsSorted();
                for (int i = 0; i < cells.Count; i++)
                {
                    SpreadsheetCell cell = cells[i];
                    if (cell == null || cell.Row < 1 || cell.Column < 1)
                        continue;
                    int rowIndex = cell.Row - 1;
                    int columnIndex = cell.Column - 1;
                    if (rowIndex < grid.Rows.Count && columnIndex < grid.Columns.Count)
                        grid.Rows[rowIndex].Cells[columnIndex].Value = cell.DisplayText;
                }
            }
            finally
            {
                loadingGrid = false;
            }

            if (grid.Rows.Count > 0 && grid.Columns.Count > 0 && grid.CurrentCell == null)
                grid.CurrentCell = grid.Rows[0].Cells[0];

            LoadCurrentCellIntoInputBar();
            UpdateStatus();
        }

        private void CommitGridCell(int rowIndex, int columnIndex)
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null || rowIndex < 0 || columnIndex < 0)
                return;

            object value = grid.Rows[rowIndex].Cells[columnIndex].Value;
            string input = value == null ? string.Empty : Convert.ToString(value, CultureInfo.CurrentCulture);
            sheet.SetInput(rowIndex + 1, columnIndex + 1, input);
            session.MarkDirty();
            LoadCurrentCellIntoInputBar();
            UpdateStatus();
        }

        private void CommitInputBar()
        {
            if (grid.CurrentCell == null)
                return;

            int rowIndex = grid.CurrentCell.RowIndex;
            int columnIndex = grid.CurrentCell.ColumnIndex;
            string input = inputBar.Text ?? string.Empty;

            loadingGrid = true;
            try
            {
                grid.CurrentCell.Value = input;
            }
            finally
            {
                loadingGrid = false;
            }

            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet != null)
            {
                sheet.SetInput(rowIndex + 1, columnIndex + 1, input);
                session.MarkDirty();
            }
            UpdateStatus();
        }

        private void LoadCurrentCellIntoInputBar()
        {
            if (grid.CurrentCell == null)
            {
                cellLabel.Text = string.Empty;
                inputBar.Text = string.Empty;
                return;
            }

            int row = grid.CurrentCell.RowIndex + 1;
            int column = grid.CurrentCell.ColumnIndex + 1;
            cellLabel.Text = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

            SpreadsheetSheet sheet = GetCurrentSheet();
            SpreadsheetCell cell = sheet == null ? null : sheet.GetCell(row, column, false);
            inputBar.Text = cell == null ? string.Empty : cell.DisplayText;
        }

        private void ClearSelectedCells()
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null)
                return;

            loadingGrid = true;
            try
            {
                for (int i = 0; i < grid.SelectedCells.Count; i++)
                {
                    DataGridViewCell selected = grid.SelectedCells[i];
                    sheet.SetInput(selected.RowIndex + 1, selected.ColumnIndex + 1, string.Empty);
                    selected.Value = null;
                }
            }
            finally
            {
                loadingGrid = false;
            }

            session.MarkDirty();
            LoadCurrentCellIntoInputBar();
            UpdateStatus();
        }

        private void AddSheet()
        {
            SpreadsheetSheet sheet = session.Document.AddSheet("Sheet");
            session.MarkDirty();
            RefreshSheetTabs();
            sheetTabs.SelectedIndex = session.Document.Sheets.IndexOf(sheet);
            UpdateStatus();
        }

        private void DeleteSheet()
        {
            int index = currentSheetIndex;
            if (!session.Document.RemoveSheet(index))
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
            currentSheetIndex = Math.Min(index, session.Document.Sheets.Count - 1);
            RefreshSheetTabs();
            sheetTabs.SelectedIndex = currentSheetIndex;
            LoadSheet();
        }

        private void RenameSheet()
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            if (sheet == null)
                return;

            string name = ViewerDialogs.Prompt(
                this,
                "Rename sheet",
                "Worksheet name:",
                sheet.Name);

            if (name == null)
                return;

            if (!session.Document.RenameSheet(currentSheetIndex, name))
            {
                MessageBox.Show(
                    this,
                    "That worksheet name is invalid or already in use.",
                    "Rename sheet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            session.MarkDirty();
            RefreshSheetTabs();
            sheetTabs.SelectedIndex = currentSheetIndex;
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
            sheetTabs.SelectedIndex = target;
            UpdateStatus();
        }

        private void RefreshSheetTabs()
        {
            int selected = Math.Max(0, Math.Min(currentSheetIndex, session.Document.Sheets.Count - 1));
            sheetTabs.TabPages.Clear();

            for (int i = 0; i < session.Document.Sheets.Count; i++)
            {
                SpreadsheetSheet sheet = session.Document.Sheets[i];
                sheetTabs.TabPages.Add(
                    new TabPage(sheet == null ? "Sheet" + (i + 1).ToString() : sheet.Name));
            }

            if (sheetTabs.TabPages.Count > 0)
                sheetTabs.SelectedIndex = selected;
        }

        private SpreadsheetSheet GetCurrentSheet()
        {
            if (currentSheetIndex < 0 || currentSheetIndex >= session.Document.Sheets.Count)
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
                        dialog.Filter = "Excel Open XML Workbook (*.xlsx)|*.xlsx";
                        dialog.DefaultExt = "xlsx";
                        dialog.AddExtension = true;
                        dialog.FileName = SafeFileName(session.Document.Title) + ".xlsx";

                        if (dialog.ShowDialog(this) != DialogResult.OK)
                            return;

                        session.SaveAs(dialog.FileName);
                    }
                }
                else
                {
                    session.Save();
                }

                Text = Path.GetFileName(session.FilePath) + " - Spreadsheet Editor";
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

        private void OnEditorClosing(object sender, FormClosingEventArgs e)
        {
            if (!session.IsDirty)
                return;

            DialogResult result = MessageBox.Show(
                this,
                "Save changes before closing?",
                "Unsaved workbook",
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

        private void UpdateStatus()
        {
            SpreadsheetSheet sheet = GetCurrentSheet();
            statusLabel.Text =
                (session.IsDirty ? "Modified  •  " : "Saved  •  ") +
                (string.IsNullOrEmpty(session.FilePath)
                    ? "Unsaved XLSX"
                    : Path.GetFileName(session.FilePath)) +
                "  •  " + session.Document.Sheets.Count.ToString() + " sheet(s)" +
                (sheet == null ? string.Empty : "  •  " + sheet.Name) +
                "  •  formulas are stored but not calculated by the editor";
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

        private static string ColumnName(int column)
        {
            int value = Math.Max(1, column);
            string result = string.Empty;
            while (value > 0)
            {
                value--;
                result = ((char)('A' + (value % 26))).ToString() + result;
                value /= 26;
            }
            return result;
        }

        private static string SafeFileName(string value)
        {
            string text = string.IsNullOrEmpty(value) ? "New Workbook" : value;
            char[] invalid = Path.GetInvalidFileNameChars();
            for (int i = 0; i < invalid.Length; i++)
                text = text.Replace(invalid[i], '_');
            return text;
        }
    }
}
