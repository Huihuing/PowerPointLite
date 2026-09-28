using System;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.Control && e.Alt && e.KeyCode == Keys.N)
            {
                ShowOfficeWorkspace();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Control && e.Alt && e.KeyCode == Keys.O)
            {
                OpenExistingEditableOfficeFile();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Control && e.Alt && e.KeyCode == Keys.D)
            {
                OpenNewDocxEditor();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Control && e.Alt && e.KeyCode == Keys.X)
            {
                OpenNewXlsxEditor();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Control && e.Alt && e.KeyCode == Keys.H)
            {
                OpenNewHwpxEditor();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Control && e.Shift && e.KeyCode == Keys.I)
            {
                ShowCurrentPresentationCompatibilityReport();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Control && e.Shift && e.KeyCode == Keys.E)
            {
                OpenCurrentPresentationEditor();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Control && e.KeyCode == Keys.N)
            {
                OpenNewPresentationEditor();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            base.OnKeyDown(e);
        }

        private void ShowOfficeWorkspace()
        {
            using (OfficeWorkspaceDialog dialog = new OfficeWorkspaceDialog())
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                if (dialog.SelectedAction == WorkspaceAction.NewPresentation)
                    OpenNewPresentationEditor();
                else if (dialog.SelectedAction == WorkspaceAction.NewDocx)
                    OpenNewDocxEditor();
                else if (dialog.SelectedAction == WorkspaceAction.NewXlsx)
                    OpenNewXlsxEditor();
                else if (dialog.SelectedAction == WorkspaceAction.NewHwpx)
                    OpenNewHwpxEditor();
                else if (dialog.SelectedAction == WorkspaceAction.OpenExisting)
                    OpenExistingEditableOfficeFile();
            }
        }

        private void OpenNewPresentationEditor()
        {
            PresentationEditSession session =
                PresentationEditSession.CreateNew("New Presentation");

            using (AdvancedPresentationEditorForm editor =
                new AdvancedPresentationEditorForm(session))
            {
                AdvancedEditorTableExtension.Attach(editor, session);
                AdvancedEditorThumbnailExtension.Attach(editor, session);
                editor.ShowDialog(this);
                LoadSavedEditorOutput(editor.SavedFilePath);
            }
        }

        private void OpenNewDocxEditor()
        {
            using (UnifiedTextDocumentEditorForm editor =
                UnifiedTextDocumentEditorForm.CreateNewDocx("New Document"))
            {
                editor.ShowDialog(this);
                ShowSavedNonPresentation(editor.SavedFilePath, "DOCX");
            }
        }

        private void OpenNewHwpxEditor()
        {
            using (UnifiedTextDocumentEditorForm editor =
                UnifiedTextDocumentEditorForm.CreateNewHwpx("New HWPX Document"))
            {
                editor.ShowDialog(this);
                ShowSavedNonPresentation(editor.SavedFilePath, "HWPX");
            }
        }

        private void OpenNewXlsxEditor()
        {
            using (SpreadsheetEditorForm editor =
                SpreadsheetEditorForm.CreateNew("New Workbook"))
            {
                editor.ShowDialog(this);
                ShowSavedNonPresentation(editor.SavedFilePath, "XLSX");
            }
        }

        private void OpenExistingEditableOfficeFile()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter =
                    "Supported documents (*.pptx;*.docx;*.xlsx;*.hwpx;*.hwp)|*.pptx;*.docx;*.xlsx;*.hwpx;*.hwp|" +
                    "PowerPoint Open XML (*.pptx)|*.pptx|" +
                    "Word Open XML (*.docx)|*.docx|" +
                    "Excel Open XML (*.xlsx)|*.xlsx|" +
                    "HWPX (*.hwpx)|*.hwpx|" +
                    "HWP 5.x read-only (*.hwp)|*.hwp|All files (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                string path = dialog.FileName;
                string extension = Path.GetExtension(path).ToLowerInvariant();

                try
                {
                    if (extension == ".pptx")
                    {
                        OpenPresentationEditorForPath(path);
                    }
                    else if (extension == ".docx" || extension == ".hwpx")
                    {
                        using (UnifiedTextDocumentEditorForm editor =
                            UnifiedTextDocumentEditorForm.Open(path))
                        {
                            editor.ShowDialog(this);
                            ShowSavedNonPresentation(
                                editor.SavedFilePath,
                                extension == ".hwpx" ? "HWPX" : "DOCX");
                        }
                    }
                    else if (extension == ".xlsx")
                    {
                        using (SpreadsheetEditorForm editor = SpreadsheetEditorForm.Open(path))
                        {
                            editor.ShowDialog(this);
                            ShowSavedNonPresentation(editor.SavedFilePath, "XLSX");
                        }
                    }
                    else if (extension == ".hwp")
                    {
                        using (HwpReadOnlyViewerForm viewerForm =
                            new HwpReadOnlyViewerForm(path))
                        {
                            viewerForm.ShowDialog(this);
                        }
                    }
                    else
                    {
                        MessageBox.Show(
                            this,
                            "The selected file is not supported by the current workspace.",
                            "Open document",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        this,
                        ex.Message +
                        "\r\n\r\nThe original file has not been simplified or overwritten. " +
                        "Experimental editors use deny-by-default safety checks, and HWP is read-only.",
                        "Document could not be opened safely",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }

        private void OpenCurrentPresentationEditor()
        {
            if (string.IsNullOrEmpty(currentFile) ||
                !File.Exists(currentFile) ||
                !string.Equals(
                    Path.GetExtension(currentFile),
                    ".pptx",
                    StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    this,
                    "Open a PPTX file first. Existing-file editing is currently limited to PPTX files that the experimental writer can round-trip safely.",
                    "Edit presentation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                OpenPresentationEditorForPath(currentFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message +
                    "\r\n\r\nThe file remains unchanged and can still be opened in Viewer mode.\r\nUse Ctrl+Shift+I for a read-only compatibility report.",
                    "Editing is not safe yet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void OpenPresentationEditorForPath(string path)
        {
            PresentationEditSession session =
                PresentationEditSession.OpenEditable(path);

            using (AdvancedPresentationEditorForm editor =
                new AdvancedPresentationEditorForm(session))
            {
                AdvancedEditorTableExtension.Attach(editor, session);
                AdvancedEditorThumbnailExtension.Attach(editor, session);
                editor.ShowDialog(this);
                LoadSavedEditorOutput(editor.SavedFilePath);
            }
        }

        private void ShowCurrentPresentationCompatibilityReport()
        {
            if (string.IsNullOrEmpty(currentFile) ||
                !File.Exists(currentFile) ||
                !string.Equals(
                    Path.GetExtension(currentFile),
                    ".pptx",
                    StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    this,
                    "Open a PPTX file first.",
                    "PPTX compatibility",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                PptxCompatibilityReport report =
                    PptxCompatibilityAnalyzer.Analyze(currentFile);

                using (PptxCompatibilityReportForm form =
                    new PptxCompatibilityReportForm(report))
                {
                    form.ShowDialog(this);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Compatibility scan failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void LoadSavedEditorOutput(string savedPath)
        {
            if (!string.IsNullOrEmpty(savedPath) &&
                File.Exists(savedPath))
            {
                LoadPresentation(savedPath);
            }
        }

        private void ShowSavedNonPresentation(string savedPath, string format)
        {
            if (string.IsNullOrEmpty(savedPath) || !File.Exists(savedPath))
                return;

            status.Text =
                format + " saved: " + Path.GetFileName(savedPath) +
                "  •  Use Ctrl+Alt+O to reopen supported documents.";
        }
    }
}
