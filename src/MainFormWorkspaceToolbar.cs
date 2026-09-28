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
            ApplicationTheme.ApplyButton(workspaceToolbarButton);
            workspaceToolbarButton.Click += delegate
            {
                ShowOfficeWorkspace();
            };
            toolbar.Controls.Add(workspaceToolbarButton);

            // Keep the original viewer controls visible while adding a direct
            // multi-format workspace entry. No Office/Hancom ribbon assets are
            // copied; this uses the project's own flat WinForms theme.
            for (int i = 0; i < toolbar.Controls.Count; i++)
            {
                Button button = toolbar.Controls[i] as Button;
                if (button != null &&
                    button != workspaceToolbarButton &&
                    button.Text.StartsWith("Auto TOC", StringComparison.Ordinal))
                {
                    button.Left = 952;
                }
            }

            zoomTrack.Left = 1042;
            engineLabel.Left = 1202;
            workspaceToolbarButton.BringToFront();
        }
    }
}
