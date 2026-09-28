using System;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        protected override void OnKeyDown(KeyEventArgs e)
        {
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
            using (PresentationEditorForm editor =
                PresentationEditorForm.CreateNew("New Presentation"))
            {
                editor.ShowDialog(this);

                string savedPath = editor.SavedFilePath;

                if (!string.IsNullOrEmpty(savedPath) &&
                    File.Exists(savedPath))
                {
                    LoadPresentation(savedPath);
                }
            }
        }
    }
}
