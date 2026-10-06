using System;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            // The document Workspace already has its own dedicated strip at
            // the bottom of the toolbar. Keep the upper row focused on Viewer
            // controls instead of showing a duplicate Workspace button.
            LayoutViewerToolbar();
        }
    }
}
