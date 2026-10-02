using System;
using System.Drawing;
using System.Windows.Forms;

namespace PptxViewer
{
    internal static class ApplicationTheme
    {
        public static readonly Color Window = Color.FromArgb(24, 26, 31);
        public static readonly Color Toolbar = Color.FromArgb(31, 34, 40);
        public static readonly Color Sidebar = Color.FromArgb(27, 30, 35);
        public static readonly Color Surface = Color.FromArgb(38, 42, 49);
        public static readonly Color SurfaceHover = Color.FromArgb(49, 54, 63);
        public static readonly Color SurfacePressed = Color.FromArgb(58, 64, 74);
        public static readonly Color Canvas = Color.FromArgb(15, 17, 20);
        public static readonly Color Divider = Color.FromArgb(52, 57, 66);
        public static readonly Color PrimaryText = Color.FromArgb(242, 244, 247);
        public static readonly Color SecondaryText = Color.FromArgb(172, 179, 189);
        public static readonly Color Accent = Color.FromArgb(91, 140, 255);
        public static readonly Color AccentSoft = Color.FromArgb(54, 78, 126);

        public static void ApplyButton(Button button)
        {
            if (button == null)
                return;

            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = Surface;
            button.ForeColor = PrimaryText;
            button.Cursor = Cursors.Hand;
            button.UseVisualStyleBackColor = false;

            button.MouseEnter += delegate
            {
                if (button.Enabled)
                    button.BackColor = SurfaceHover;
            };

            button.MouseLeave += delegate
            {
                if (button.Enabled)
                    button.BackColor = Surface;
            };

            button.MouseDown += delegate(object sender, MouseEventArgs e)
            {
                if (button.Enabled && e.Button == MouseButtons.Left)
                    button.BackColor = SurfacePressed;
            };

            button.MouseUp += delegate
            {
                if (button.Enabled)
                    button.BackColor = SurfaceHover;
            };
        }

        public static void ApplyContextMenu(ContextMenuStrip menu)
        {
            if (menu == null)
                return;

            menu.BackColor = Toolbar;
            menu.ForeColor = PrimaryText;
            menu.ShowImageMargin = false;

            for (int i = 0; i < menu.Items.Count; i++)
            {
                ToolStripItem item = menu.Items[i];
                item.BackColor = Toolbar;
                item.ForeColor = PrimaryText;
            }
        }

        public static void ApplyRecursively(Control control)
        {
            if (control == null)
                return;

            Button button = control as Button;
            if (button != null)
            {
                ApplyButton(button);
            }
            else if (control is Label)
            {
                control.ForeColor = PrimaryText;
            }
            else if (control is FlowLayoutPanel)
            {
                control.BackColor = Sidebar;
            }
            else if (control is Panel)
            {
                control.BackColor = Surface;
            }

            for (int i = 0; i < control.Controls.Count; i++)
                ApplyRecursively(control.Controls[i]);
        }
    }

    public sealed partial class MainForm : Form
    {
        private bool uiThemeApplied;

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            if (uiThemeApplied)
                return;

            uiThemeApplied = true;
            ApplyIndependentUiTheme();
        }

        private void ApplyIndependentUiTheme()
        {
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;

            try
            {
                Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            }
            catch
            {
                Font = SystemFonts.MessageBoxFont;
            }

            toolbar.Height = 84;
            toolbar.BackColor = ApplicationTheme.Toolbar;

            mainSplit.BackColor = ApplicationTheme.Divider;
            thumbnails.BackColor = ApplicationTheme.Sidebar;
            viewerPanel.BackColor = ApplicationTheme.Canvas;
            viewer.BackColor = Color.Black;

            status.Height = 28;
            status.BackColor = ApplicationTheme.Toolbar;
            status.ForeColor = ApplicationTheme.SecondaryText;
            status.TextAlign = ContentAlignment.MiddleLeft;
            status.Padding = new Padding(12, 0, 8, 0);

            engineLabel.ForeColor = ApplicationTheme.SecondaryText;

            notesBox.BorderStyle = BorderStyle.None;
            notesBox.BackColor = Color.FromArgb(246, 247, 249);
            notesBox.ForeColor = Color.FromArgb(30, 33, 38);

            for (int i = 0; i < toolbar.Controls.Count; i++)
                ApplicationTheme.ApplyRecursively(toolbar.Controls[i]);

            // Keep the toolbar itself visually distinct from ordinary surfaces.
            toolbar.BackColor = ApplicationTheme.Toolbar;

            AddDocumentWorkspaceStrip();
            ApplicationTheme.ApplyContextMenu(slideshowMenu);
            AdvancedEditorFidelityExtension.Initialize();
            InitializePptxFidelityFallback();

            thumbnails.ControlAdded += delegate(object sender, ControlEventArgs e)
            {
                ApplicationTheme.ApplyRecursively(e.Control);
            };

            // The theme is intentionally original and asset-free. It does not
            // copy Office icons, ribbons, artwork, or proprietary UI resources.
        }

        private void AddDocumentWorkspaceStrip()
        {
            if (toolbar.Controls.ContainsKey("DocumentWorkspaceStrip"))
                return;

            Panel strip = new Panel();
            strip.Name = "DocumentWorkspaceStrip";
            strip.Dock = DockStyle.Bottom;
            strip.Height = 34;
            strip.BackColor = Color.FromArgb(26, 29, 34);
            strip.Padding = new Padding(8, 2, 8, 2);

            Label section = new Label();
            section.AutoSize = false;
            section.Left = 10;
            section.Top = 7;
            section.Width = 82;
            section.Height = 20;
            section.Text = "Documents";
            section.Font = new Font(Font, FontStyle.Bold);
            section.ForeColor = ApplicationTheme.PrimaryText;
            strip.Controls.Add(section);

            Button workspace = new Button();
            workspace.Left = 96;
            workspace.Top = 3;
            workspace.Width = 106;
            workspace.Height = 28;
            workspace.Text = "New / Open";
            workspace.TabStop = true;
            ApplicationTheme.ApplyButton(workspace);
            workspace.Click += delegate { ShowOfficeWorkspace(); };
            strip.Controls.Add(workspace);

            Label formats = new Label();
            formats.Left = 216;
            formats.Top = 7;
            formats.Width = 520;
            formats.Height = 20;
            formats.Text = "PPTX  ·  DOCX  ·  XLSX  ·  HWPX experimental  ·  HWP read-only";
            formats.ForeColor = ApplicationTheme.SecondaryText;
            strip.Controls.Add(formats);

            Label shortcut = new Label();
            shortcut.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            shortcut.Top = 7;
            shortcut.Width = 210;
            shortcut.Height = 20;
            shortcut.Left = Math.Max(746, toolbar.ClientSize.Width - shortcut.Width - 12);
            shortcut.TextAlign = ContentAlignment.MiddleRight;
            shortcut.Text = "Workspace  Ctrl+Alt+N";
            shortcut.ForeColor = ApplicationTheme.SecondaryText;
            strip.Controls.Add(shortcut);

            strip.Resize += delegate
            {
                shortcut.Left = Math.Max(
                    746,
                    strip.ClientSize.Width - shortcut.Width - 12);
            };

            toolbar.Controls.Add(strip);
            strip.BringToFront();
        }
    }
}
