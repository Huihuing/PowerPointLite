using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class UnifiedTextDocumentEditorForm : Form
    {
        private readonly TextDocumentEditSession session;
        private readonly ListBox paragraphList;
        private readonly ComboBox runList;
        private readonly RichTextBox preview;
        private readonly TextBox runText;
        private readonly ComboBox fontPicker;
        private readonly NumericUpDown fontSize;
        private readonly CheckBox boldCheck;
        private readonly CheckBox italicCheck;
        private readonly CheckBox underlineCheck;
        private readonly ComboBox alignmentPicker;
        private readonly TextBox colorEditor;
        private readonly Label formatBadge;
        private readonly Label statusLabel;
        private bool loading;

        public string SavedFilePath
        {
            get { return session.FilePath; }
        }

        public UnifiedTextDocumentEditorForm(TextDocumentEditSession editSession)
        {
            if (editSession == null)
                throw new ArgumentNullException("editSession");

            session = editSession;

            Text = session.FormatDisplayName + " Document Editor";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1280;
            Height = 820;
            MinimumSize = new Size(940, 620);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);
            KeyPreview = true;

            Panel toolbar = new Panel();
            toolbar.Dock = DockStyle.Top;
            toolbar.Height = 50;
            toolbar.Padding = new Padding(8);
            toolbar.BackColor = ApplicationTheme.Toolbar;
            Controls.Add(toolbar);

            formatBadge = new Label();
            formatBadge.Left = 8;
            formatBadge.Top = 10;
            formatBadge.Width = 66;
            formatBadge.Height = 30;
            formatBadge.Text = session.FormatDisplayName;
            formatBadge.TextAlign = ContentAlignment.MiddleCenter;
            formatBadge.BackColor = ApplicationTheme.AccentSoft;
            formatBadge.ForeColor = ApplicationTheme.PrimaryText;
            formatBadge.Font = new Font(Font, FontStyle.Bold);
            toolbar.Controls.Add(formatBadge);

            Button save = MakeButton("Save", 84, 64);
            Button saveAs = MakeButton("Save As", 154, 76);
            Button addParagraph = MakeButton("+ Paragraph", 240, 94);
            Button deleteParagraph = MakeButton("Delete", 340, 68);
            Button moveUp = MakeButton("Up", 414, 52);
            Button moveDown = MakeButton("Down", 472, 62);
            Button addRun = MakeButton("+ Run", 544, 62);
            Button deleteRun = MakeButton("- Run", 612, 62);

            toolbar.Controls.Add(save);
            toolbar.Controls.Add(saveAs);
            toolbar.Controls.Add(addParagraph);
            toolbar.Controls.Add(deleteParagraph);
            toolbar.Controls.Add(moveUp);
            toolbar.Controls.Add(moveDown);
            toolbar.Controls.Add(addRun);
            toolbar.Controls.Add(deleteRun);

            save.Click += delegate { SaveDocument(false); };
            saveAs.Click += delegate { SaveDocument(true); };
            addParagraph.Click += delegate { AddParagraph(); };
            deleteParagraph.Click += delegate { DeleteParagraph(); };
            moveUp.Click += delegate { MoveParagraph(-1); };
            moveDown.Click += delegate { MoveParagraph(1); };
            addRun.Click += delegate { AddRun(); };
            deleteRun.Click += delegate { DeleteRun(); };

            statusLabel = new Label();
            statusLabel.Dock = DockStyle.Bottom;
            statusLabel.Height = 28;
            statusLabel.Padding = new Padding(10, 0, 10, 0);
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusLabel.BackColor = ApplicationTheme.Toolbar;
            statusLabel.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(statusLabel);

            Panel left = new Panel();
            left.Dock = DockStyle.Left;
            left.Width = 230;
            left.Padding = new Padding(10);
            left.BackColor = ApplicationTheme.Sidebar;
            Controls.Add(left);

            Label paragraphTitle = MakeLabel("Paragraphs", 0, 0, 200, 28, true);
            paragraphTitle.Dock = DockStyle.Top;
            left.Controls.Add(paragraphTitle);

            paragraphList = new ListBox();
            paragraphList.Dock = DockStyle.Fill;
            paragraphList.BorderStyle = BorderStyle.None;
            paragraphList.BackColor = ApplicationTheme.Sidebar;
            paragraphList.ForeColor = ApplicationTheme.PrimaryText;
            paragraphList.IntegralHeight = false;
            left.Controls.Add(paragraphList);
            paragraphList.BringToFront();

            Panel right = new Panel();
            right.Dock = DockStyle.Right;
            right.Width = 310;
            right.Padding = new Padding(14);
            right.BackColor = ApplicationTheme.Sidebar;
            Controls.Add(right);

            right.Controls.Add(MakeLabel("Text properties", 0, 0, 270, 26, true));
            right.Controls.Add(MakeLabel("Run", 0, 34, 80, 20, false));

            runList = new ComboBox();
            runList.Left = 0;
            runList.Top = 56;
            runList.Width = 280;
            runList.DropDownStyle = ComboBoxStyle.DropDownList;
            right.Controls.Add(runList);

            right.Controls.Add(MakeLabel("Text", 0, 90, 80, 20, false));

            runText = new TextBox();
            runText.Left = 0;
            runText.Top = 112;
            runText.Width = 280;
            runText.Height = 116;
            runText.Multiline = true;
            runText.ScrollBars = ScrollBars.Vertical;
            runText.BackColor = Color.FromArgb(245, 246, 248);
            runText.ForeColor = Color.FromArgb(28, 31, 36);
            runText.BorderStyle = BorderStyle.FixedSingle;
            right.Controls.Add(runText);

            right.Controls.Add(MakeLabel("Font", 0, 238, 80, 20, false));

            fontPicker = new ComboBox();
            fontPicker.Left = 0;
            fontPicker.Top = 260;
            fontPicker.Width = 280;
            fontPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            fontPicker.BackColor = Color.White;
            fontPicker.ForeColor = Color.FromArgb(28, 31, 36);
            right.Controls.Add(fontPicker);

            List<string> fonts = SystemFontCatalog.GetInstalledFamilyNames();
            for (int i = 0; i < fonts.Count; i++)
                fontPicker.Items.Add(fonts[i]);

            right.Controls.Add(MakeLabel("Size", 0, 294, 80, 20, false));

            fontSize = new NumericUpDown();
            fontSize.Left = 0;
            fontSize.Top = 316;
            fontSize.Width = 82;
            fontSize.Minimum = 1;
            fontSize.Maximum = 400;
            fontSize.DecimalPlaces = 1;
            fontSize.Increment = 0.5M;
            right.Controls.Add(fontSize);

            boldCheck = MakeCheck("Bold", 96, 318, 62);
            italicCheck = MakeCheck("Italic", 160, 318, 62);
            underlineCheck = MakeCheck("Underline", 222, 318, 78);
            right.Controls.Add(boldCheck);
            right.Controls.Add(italicCheck);
            right.Controls.Add(underlineCheck);

            right.Controls.Add(MakeLabel("Alignment", 0, 356, 100, 20, false));
            alignmentPicker = new ComboBox();
            alignmentPicker.Left = 0;
            alignmentPicker.Top = 378;
            alignmentPicker.Width = 132;
            alignmentPicker.DropDownStyle = ComboBoxStyle.DropDownList;
            alignmentPicker.Items.Add("Left");
            alignmentPicker.Items.Add("Center");
            alignmentPicker.Items.Add("Right");
            alignmentPicker.Items.Add("Justify");
            right.Controls.Add(alignmentPicker);

            right.Controls.Add(MakeLabel("Text color", 148, 356, 100, 20, false));
            colorEditor = new TextBox();
            colorEditor.Left = 148;
            colorEditor.Top = 378;
            colorEditor.Width = 132;
            colorEditor.MaxLength = 7;
            right.Controls.Add(colorEditor);

            Label licensing = MakeLabel(
                "Fonts come from Windows-installed families. Only the family name is written; font files are not automatically bundled or embedded.",
                0,
                428,
                280,
                88,
                false);
            licensing.ForeColor = ApplicationTheme.SecondaryText;
            right.Controls.Add(licensing);

            if (session.Format == TextDocumentFileFormat.Hwpx)
            {
                Label experimental = MakeLabel(
                    "HWPX writing is experimental until generated files are verified in a real Hancom environment.",
                    0,
                    522,
                    280,
                    66,
                    false);
                experimental.ForeColor = Color.FromArgb(238, 190, 94);
                right.Controls.Add(experimental);
            }

            Panel host = new Panel();
            host.Dock = DockStyle.Fill;
            host.Padding = new Padding(34, 28, 34, 28);
            host.BackColor = ApplicationTheme.Canvas;
            Controls.Add(host);
            host.BringToFront();

            preview = new RichTextBox();
            preview.Dock = DockStyle.Fill;
            preview.ReadOnly = true;
            preview.BorderStyle = BorderStyle.FixedSingle;
            preview.BackColor = Color.White;
            preview.ForeColor = Color.FromArgb(32, 36, 42);
            preview.DetectUrls = true;
            host.Controls.Add(preview);

            paragraphList.SelectedIndexChanged += delegate
            {
                RefreshRunList();
                LoadProperties();
                RenderPreview();
            };
            runList.SelectedIndexChanged += delegate { LoadProperties(); };

            runText.TextChanged += delegate { ApplyProperties(); };
            fontPicker.SelectedIndexChanged += delegate { ApplyProperties(); };
            fontSize.ValueChanged += delegate { ApplyProperties(); };
            boldCheck.CheckedChanged += delegate { ApplyProperties(); };
            italicCheck.CheckedChanged += delegate { ApplyProperties(); };
            underlineCheck.CheckedChanged += delegate { ApplyProperties(); };
            alignmentPicker.SelectedIndexChanged += delegate { ApplyProperties(); };
            colorEditor.TextChanged += delegate { ApplyProperties(); };

            FormClosing += OnEditorClosing;

            RefreshParagraphList();
            if (paragraphList.Items.Count > 0)
                paragraphList.SelectedIndex = 0;
            RenderPreview();
            UpdateStatus();
        }

        public static UnifiedTextDocumentEditorForm CreateNewDocx(string title)
        {
            return new UnifiedTextDocumentEditorForm(
                TextDocumentEditSession.CreateNew(title, TextDocumentFileFormat.Docx));
        }

        public static UnifiedTextDocumentEditorForm CreateNewHwpx(string title)
        {
            return new UnifiedTextDocumentEditorForm(
                TextDocumentEditSession.CreateNew(title, TextDocumentFileFormat.Hwpx));
        }

        public static UnifiedTextDocumentEditorForm Open(string path)
        {
            return new UnifiedTextDocumentEditorForm(
                TextDocumentEditSession.Open(path));
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
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void AddParagraph()
        {
            session.AddParagraph(string.Empty);
            RefreshParagraphList();
            paragraphList.SelectedIndex = session.Document.Paragraphs.Count - 1;
            UpdateStatus();
        }

        private void DeleteParagraph()
        {
            int index = paragraphList.SelectedIndex;
            if (index < 0)
                return;

            if (!session.RemoveParagraph(index))
            {
                MessageBox.Show(
                    this,
                    "A document must keep at least one paragraph.",
                    "Delete paragraph",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            RefreshParagraphList();
            paragraphList.SelectedIndex = Math.Min(index, paragraphList.Items.Count - 1);
            UpdateStatus();
        }

        private void MoveParagraph(int direction)
        {
            int from = paragraphList.SelectedIndex;
            int to = from + direction;
            if (!session.MoveParagraph(from, to))
                return;
            RefreshParagraphList();
            paragraphList.SelectedIndex = to;
            UpdateStatus();
        }

        private void AddRun()
        {
            DocumentParagraph paragraph = GetSelectedParagraph();
            if (paragraph == null)
                return;
            paragraph.AddRun(string.Empty);
            session.MarkDirty();
            RefreshRunList();
            runList.SelectedIndex = paragraph.Runs.Count - 1;
            UpdateStatus();
        }

        private void DeleteRun()
        {
            DocumentParagraph paragraph = GetSelectedParagraph();
            int index = runList.SelectedIndex;
            if (paragraph == null || index < 0 || index >= paragraph.Runs.Count)
                return;
            paragraph.Runs.RemoveAt(index);
            if (paragraph.Runs.Count == 0)
                paragraph.AddRun(string.Empty);
            session.MarkDirty();
            RefreshRunList();
            runList.SelectedIndex = Math.Min(index, paragraph.Runs.Count - 1);
            RenderPreview();
            UpdateStatus();
        }

        private void RefreshParagraphList()
        {
            int selected = paragraphList.SelectedIndex;
            paragraphList.BeginUpdate();
            paragraphList.Items.Clear();
            for (int i = 0; i < session.Document.Paragraphs.Count; i++)
            {
                DocumentParagraph paragraph = session.Document.Paragraphs[i];
                string text = paragraph == null ? string.Empty : FirstLine(paragraph.PlainText);
                if (string.IsNullOrEmpty(text))
                    text = "(empty paragraph)";
                paragraphList.Items.Add((i + 1).ToString() + ".  " + text);
            }
            paragraphList.EndUpdate();
            if (paragraphList.Items.Count > 0 && selected >= 0)
                paragraphList.SelectedIndex = Math.Min(selected, paragraphList.Items.Count - 1);
        }

        private void RefreshRunList()
        {
            int selected = runList.SelectedIndex;
            runList.Items.Clear();
            DocumentParagraph paragraph = GetSelectedParagraph();
            if (paragraph == null)
                return;

            if (paragraph.Runs.Count == 0)
            {
                paragraph.AddRun(string.Empty);
                session.MarkDirty();
            }

            for (int i = 0; i < paragraph.Runs.Count; i++)
            {
                DocumentTextRun run = paragraph.Runs[i];
                string text = run == null ? string.Empty : FirstLine(run.Text);
                if (string.IsNullOrEmpty(text))
                    text = "(empty run)";
                runList.Items.Add((i + 1).ToString() + ". " + text);
            }

            runList.SelectedIndex =
                selected >= 0 && selected < runList.Items.Count ? selected : 0;
        }

        private void LoadProperties()
        {
            loading = true;
            try
            {
                DocumentParagraph paragraph = GetSelectedParagraph();
                DocumentTextRun run = GetSelectedRun();
                bool enabled = run != null;

                runText.Enabled = enabled;
                fontPicker.Enabled = enabled;
                fontSize.Enabled = enabled;
                boldCheck.Enabled = enabled;
                italicCheck.Enabled = enabled;
                underlineCheck.Enabled = enabled;
                colorEditor.Enabled = enabled;
                alignmentPicker.Enabled = paragraph != null;

                alignmentPicker.SelectedIndex = paragraph == null ? -1 : (int)paragraph.Alignment;

                if (!enabled)
                {
                    runText.Text = string.Empty;
                    colorEditor.Text = string.Empty;
                    return;
                }

                runText.Text = run.Text ?? string.Empty;
                SelectFont(run.FontFamily);
                fontSize.Value = (decimal)Math.Max(1f, Math.Min(400f, run.FontSizePoints));
                boldCheck.Checked = run.Bold;
                italicCheck.Checked = run.Italic;
                underlineCheck.Checked = run.Underline;
                colorEditor.Text = "#" + NormalizeHex(run.ColorHex);
            }
            finally
            {
                loading = false;
            }
        }

        private void ApplyProperties()
        {
            if (loading)
                return;

            DocumentParagraph paragraph = GetSelectedParagraph();
            DocumentTextRun run = GetSelectedRun();

            if (paragraph != null && alignmentPicker.SelectedIndex >= 0 && alignmentPicker.SelectedIndex <= 3)
                paragraph.Alignment = (DocumentParagraphAlignment)alignmentPicker.SelectedIndex;

            if (run != null)
            {
                run.Text = runText.Text ?? string.Empty;
                if (fontPicker.SelectedItem != null)
                    run.FontFamily = fontPicker.SelectedItem.ToString();
                run.FontSizePoints = (float)fontSize.Value;
                run.Bold = boldCheck.Checked;
                run.Italic = italicCheck.Checked;
                run.Underline = underlineCheck.Checked;
                string color = NormalizeHex(colorEditor.Text);
                if (color.Length == 6)
                    run.ColorHex = color;
            }

            session.MarkDirty();
            RefreshSelectedLabels();
            RenderPreview();
            UpdateStatus();
        }

        private void RefreshSelectedLabels()
        {
            int p = paragraphList.SelectedIndex;
            DocumentParagraph paragraph = GetSelectedParagraph();
            if (p >= 0 && paragraph != null)
            {
                string text = FirstLine(paragraph.PlainText);
                if (string.IsNullOrEmpty(text))
                    text = "(empty paragraph)";
                paragraphList.Items[p] = (p + 1).ToString() + ".  " + text;
            }

            int r = runList.SelectedIndex;
            DocumentTextRun run = GetSelectedRun();
            if (r >= 0 && run != null)
            {
                string text = FirstLine(run.Text);
                if (string.IsNullOrEmpty(text))
                    text = "(empty run)";
                runList.Items[r] = (r + 1).ToString() + ". " + text;
            }
        }

        private void RenderPreview()
        {
            preview.SuspendLayout();
            preview.Clear();

            for (int p = 0; p < session.Document.Paragraphs.Count; p++)
            {
                DocumentParagraph paragraph = session.Document.Paragraphs[p];
                int paragraphStart = preview.TextLength;

                if (paragraph != null)
                {
                    for (int r = 0; r < paragraph.Runs.Count; r++)
                    {
                        DocumentTextRun run = paragraph.Runs[r];
                        if (run == null)
                            continue;
                        int start = preview.TextLength;
                        preview.AppendText(run.Text ?? string.Empty);
                        int length = preview.TextLength - start;
                        if (length > 0)
                        {
                            preview.Select(start, length);
                            preview.SelectionFont = SafeFont(run);
                            preview.SelectionColor = ParseColor(run.ColorHex, Color.FromArgb(32, 36, 42));
                        }
                    }
                }

                int paragraphEnd = preview.TextLength;
                preview.AppendText(Environment.NewLine);
                preview.Select(paragraphStart, Math.Max(0, paragraphEnd - paragraphStart));
                if (paragraph != null)
                {
                    preview.SelectionAlignment =
                        paragraph.Alignment == DocumentParagraphAlignment.Center
                            ? HorizontalAlignment.Center
                            : paragraph.Alignment == DocumentParagraphAlignment.Right
                                ? HorizontalAlignment.Right
                                : HorizontalAlignment.Left;
                }
            }

            preview.Select(0, 0);
            preview.ResumeLayout();
        }

        private void SaveDocument(bool saveAs)
        {
            try
            {
                if (saveAs || string.IsNullOrEmpty(session.FilePath))
                {
                    using (SaveFileDialog dialog = new SaveFileDialog())
                    {
                        dialog.Filter = session.SaveDialogFilter;
                        dialog.DefaultExt = session.DefaultExtension;
                        dialog.AddExtension = true;
                        dialog.FileName = SafeFileName(session.Document.Title) + "." + session.DefaultExtension;
                        if (dialog.ShowDialog(this) != DialogResult.OK)
                            return;
                        session.SaveAs(dialog.FileName);
                    }
                }
                else
                {
                    session.Save();
                }

                Text = Path.GetFileName(session.FilePath) + " - " + session.FormatDisplayName + " Editor";
                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Save failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnEditorClosing(object sender, FormClosingEventArgs e)
        {
            if (!session.IsDirty)
                return;

            DialogResult result = MessageBox.Show(
                this,
                "Save changes before closing?",
                "Unsaved " + session.FormatDisplayName + " document",
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

        private DocumentParagraph GetSelectedParagraph()
        {
            int index = paragraphList.SelectedIndex;
            if (index < 0 || index >= session.Document.Paragraphs.Count)
                return null;
            return session.Document.Paragraphs[index];
        }

        private DocumentTextRun GetSelectedRun()
        {
            DocumentParagraph paragraph = GetSelectedParagraph();
            int index = runList.SelectedIndex;
            if (paragraph == null || index < 0 || index >= paragraph.Runs.Count)
                return null;
            return paragraph.Runs[index];
        }

        private void SelectFont(string family)
        {
            if (string.IsNullOrEmpty(family))
                return;
            for (int i = 0; i < fontPicker.Items.Count; i++)
            {
                if (string.Equals(fontPicker.Items[i].ToString(), family, StringComparison.CurrentCultureIgnoreCase))
                {
                    fontPicker.SelectedIndex = i;
                    return;
                }
            }
        }

        private void UpdateStatus()
        {
            statusLabel.Text =
                (session.IsDirty ? "Modified  •  " : "Saved  •  ") +
                (string.IsNullOrEmpty(session.FilePath)
                    ? "Unsaved " + session.FormatDisplayName
                    : Path.GetFileName(session.FilePath)) +
                "  •  " + session.Document.Paragraphs.Count.ToString() + " paragraph(s)";
        }

        private Button MakeButton(string text, int left, int width)
        {
            Button button = new Button();
            button.Text = text;
            button.Left = left;
            button.Top = 9;
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

        private CheckBox MakeCheck(string text, int left, int top, int width)
        {
            CheckBox check = new CheckBox();
            check.Text = text;
            check.Left = left;
            check.Top = top;
            check.Width = width;
            check.ForeColor = ApplicationTheme.PrimaryText;
            return check;
        }

        private Font SafeFont(DocumentTextRun run)
        {
            FontStyle style = FontStyle.Regular;
            if (run.Bold) style |= FontStyle.Bold;
            if (run.Italic) style |= FontStyle.Italic;
            if (run.Underline) style |= FontStyle.Underline;
            try
            {
                return new Font(
                    string.IsNullOrEmpty(run.FontFamily) ? "Arial" : run.FontFamily,
                    Math.Max(1f, Math.Min(96f, run.FontSizePoints)),
                    style,
                    GraphicsUnit.Point);
            }
            catch
            {
                return new Font(
                    SystemFonts.MessageBoxFont.FontFamily,
                    Math.Max(1f, Math.Min(96f, run.FontSizePoints)),
                    style,
                    GraphicsUnit.Point);
            }
        }

        private static string FirstLine(string value)
        {
            string text = value ?? string.Empty;
            int newline = text.IndexOfAny(new char[] { '\r', '\n' });
            if (newline >= 0)
                text = text.Substring(0, newline);
            if (text.Length > 42)
                text = text.Substring(0, 42) + "…";
            return text;
        }

        private static string SafeFileName(string value)
        {
            string text = string.IsNullOrEmpty(value) ? "New Document" : value;
            char[] invalid = Path.GetInvalidFileNameChars();
            for (int i = 0; i < invalid.Length; i++)
                text = text.Replace(invalid[i], '_');
            return text;
        }

        private static string NormalizeHex(string value)
        {
            string candidate = (value ?? string.Empty).Trim().TrimStart('#').ToUpperInvariant();
            int parsed;
            if (candidate.Length == 6 && int.TryParse(candidate, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out parsed))
                return candidate;
            return string.Empty;
        }

        private static Color ParseColor(string value, Color fallback)
        {
            string candidate = NormalizeHex(value);
            int parsed;
            if (candidate.Length == 6 && int.TryParse(candidate, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out parsed))
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
