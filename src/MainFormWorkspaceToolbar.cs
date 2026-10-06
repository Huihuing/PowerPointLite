using System;
using System.Drawing;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        private Button workspaceToolbarButton;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            if (workspaceToolbarButton != null)
                return;

            workspaceToolbarButton = new Button();
            workspaceToolbarButton.Text = "Workspace";
            workspaceToolbarButton.Left = 862;
            workspaceToolbarButton.Top = 8;
            workspaceToolbarButton.Width = 84;
            workspaceToolbarButton.Height = 30;
            workspaceToolbarButton.Tag = "Workspace";
            workspaceToolbarButton.TextAlign = ContentAlignment.MiddleCenter;
            workspaceToolbarButton.UseCompatibleTextRendering = true;
            workspaceToolbarButton.Padding = new Padding(0, 1, 0, 0);
            ApplicationTheme.ApplyButton(workspaceToolbarButton);
            workspaceToolbarButton.Click += delegate
            {
                ShowOfficeWorkspace();
            };
            toolbar.Controls.Add(workspaceToolbarButton);

            // Reflow the viewer toolbar after adding Workspace. The layout
            // helper hides the redundant Auto TOC toolbar button while the
            // same function remains available from the View menu.
            LayoutViewerToolbar();
            workspaceToolbarButton.BringToFront();
        }
    }
}
