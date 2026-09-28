using System;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        protected override void OnKeyDown(KeyEventArgs e)
        {
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
                PresentationEditSession session =
                    PresentationEditSession.OpenEditable(currentFile);

                using (AdvancedPresentationEditorForm editor =
                    new AdvancedPresentationEditorForm(session))
                {
                    AdvancedEditorTableExtension.Attach(editor, session);
                    AdvancedEditorThumbnailExtension.Attach(editor, session);
                    editor.ShowDialog(this);
                    LoadSavedEditorOutput(editor.SavedFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message +
                    "\r\n\r\nThe file remains unchanged and can still be opened in Viewer mode.",
                    "Editing is not safe yet",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
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
    }
}
