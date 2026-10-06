using System;
using System.Drawing;
using System.Windows.Forms;

namespace PptxViewer
{
    internal static class ApplicationTheme
    {
        public static readonly Color Window = Color.FromArgb(23, 25, 29);
        public static readonly Color Toolbar = Color.FromArgb(29, 32, 37);
        public static readonly Color Sidebar = Color.FromArgb(26, 29, 34);
        public static readonly Color Surface = Color.FromArgb(37, 41, 47);
        public static readonly Color SurfaceHover = Color.FromArgb(47, 52, 60);
        public static readonly Color SurfacePressed = Color.FromArgb(55, 61, 71);
        public static readonly Color Canvas = Color.FromArgb(31, 34, 39);
        public static readonly Color CanvasEdge = Color.FromArgb(78, 84, 95);
        public static readonly Color CanvasShadow = Color.FromArgb(82, 0, 0, 0);
        public static readonly Color Divider = Color.FromArgb(57, 62, 72);
        public static readonly Color PrimaryText = Color.FromArgb(243, 245, 248);
        public static readonly Color SecondaryText = Color.FromArgb(166, 173, 184);
        public static readonly Color Accent = Color.FromArgb(76, 125, 255);
        public static readonly Color AccentHover = Color.FromArgb(92, 139, 255);
        public static readonly Color AccentPressed = Color.FromArgb(61, 108, 228);
        public static readonly Color AccentSoft = Color.FromArgb(48, 68, 107);
        public static readonly Color AccentSoftHover = Color.FromArgb(60, 84, 132);
        public static readonly Color AccentSoftPressed = Color.FromArgb(43, 61, 96);

        public static void ApplyButton(Button button)
        {
            if (button == null)
                return;

            Color baseColor = Surface;
            Color hoverColor = SurfaceHover;
            Color pressedColor = SurfacePressed;

            string role =
                button.Tag as string ??
                string.Empty;

            if (string.Equals(
                    role,
                    "Open",
                    StringComparison.Ordinal))
            {
                baseColor = Accent;
                hoverColor = AccentHover;
                pressedColor = AccentPressed;
            }
            else if (string.Equals(
                         role,
                         "F5 Show",
                         StringComparison.Ordinal))
            {
                baseColor = AccentSoft;
                hoverColor = AccentSoftHover;
                pressedColor = AccentSoftPressed;
            }

            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = hoverColor;
            button.FlatAppearance.MouseDownBackColor = pressedColor;
            button.BackColor = baseColor;
            button.ForeColor = PrimaryText;
            button.Cursor = Cursors.Hand;
            button.UseVisualStyleBackColor = false;
            button.TextAlign = ContentAlignment.MiddleCenter;
            button.UseCompatibleTextRendering = true;
            button.Padding = new Padding(8, 0, 8, 1);
            button.AutoEllipsis = true;
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
                Font = new Font("Segoe UI", 9.25f, FontStyle.Regular, GraphicsUnit.Point);
            }
            catch
            {
                Font = SystemFonts.MessageBoxFont;
            }

            toolbar.Height = 82;
            toolbar.BackColor = ApplicationTheme.Toolbar;

            mainSplit.BackColor = ApplicationTheme.Divider;
            thumbnails.BackColor = ApplicationTheme.Sidebar;
            viewerPanel.BackColor = ApplicationTheme.Canvas;
            viewer.BackColor = Color.FromArgb(8, 9, 11);
            viewer.BorderStyle = BorderStyle.None;

            viewerPanel.Paint += delegate(object sender, PaintEventArgs e)
            {
                if (viewer == null ||
                    viewer.Image == null ||
                    viewer.Width <= 0 ||
                    viewer.Height <= 0)
                {
                    return;
                }

                Rectangle slideBounds = viewer.Bounds;
                Rectangle shadowBounds = slideBounds;
                shadowBounds.Inflate(5, 5);
                shadowBounds.Offset(2, 3);

                using (Pen shadow = new Pen(
                    ApplicationTheme.CanvasShadow,
                    4f))
                using (Pen edge = new Pen(
                    ApplicationTheme.CanvasEdge,
                    1f))
                {
                    e.Graphics.DrawRectangle(
                        shadow,
                        shadowBounds);
                    e.Graphics.DrawRectangle(
                        edge,
                        new Rectangle(
                            slideBounds.Left - 1,
                            slideBounds.Top - 1,
                            slideBounds.Width + 1,
                            slideBounds.Height + 1));
                }
            };

            EventHandler invalidateCanvas = delegate
            {
                viewerPanel.Invalidate();
            };

            viewer.LocationChanged += invalidateCanvas;
            viewer.SizeChanged += invalidateCanvas;
            viewer.ImageChanged += invalidateCanvas;
            viewerPanel.Resize += invalidateCanvas;

            status.Height = 30;
            status.BackColor = ApplicationTheme.Toolbar;
            status.ForeColor = ApplicationTheme.SecondaryText;
            status.TextAlign = ContentAlignment.MiddleLeft;
            status.Padding = new Padding(14, 0, 12, 0);
            status.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (Pen divider = new Pen(ApplicationTheme.Divider))
                {
                    e.Graphics.DrawLine(
                        divider,
                        0,
                        0,
                        Math.Max(0, status.ClientSize.Width - 1),
                        0);
                }
            };

            engineLabel.ForeColor = ApplicationTheme.SecondaryText;

            notesBox.BorderStyle = BorderStyle.None;
            notesBox.BackColor = Color.FromArgb(246, 247, 249);
            notesBox.ForeColor = Color.FromArgb(30, 33, 38);

            for (int i = 0; i < toolbar.Controls.Count; i++)
                ApplicationTheme.ApplyRecursively(toolbar.Controls[i]);

            toolbar.BackColor = ApplicationTheme.Toolbar;

            AddDocumentWorkspaceStrip();
            ApplicationTheme.ApplyContextMenu(slideshowMenu);
            AdvancedEditorFidelityExtension.Initialize();
            InitializePptxFidelityFallback();

            thumbnails.ControlAdded += delegate(object sender, ControlEventArgs e)
            {
                ApplicationTheme.ApplyRecursively(e.Control);
            };

            // Original asset-free UI only. No Office/Hancom artwork or ribbon assets.
        }

        private void AddDocumentWorkspaceStrip()
        {
            if (toolbar.Controls.ContainsKey("DocumentWorkspaceStrip"))
                return;

            Panel strip = new Panel();
            strip.Name = "DocumentWorkspaceStrip";
            strip.Dock = DockStyle.Bottom;
            strip.Height = 32;
            strip.BackColor = Color.FromArgb(25, 28, 33);
            strip.Padding = new Padding(10, 2, 10, 2);
            strip.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (Pen divider = new Pen(ApplicationTheme.Divider))
                {
                    e.Graphics.DrawLine(
                        divider,
                        0,
                        0,
                        Math.Max(0, strip.ClientSize.Width - 1),
                        0);
                }
            };

            Label section = new Label();
            section.AutoSize = false;
            section.Left = 10;
            section.Top = 6;
            section.Width = 82;
            section.Height = 20;
            section.Font = new Font(Font, FontStyle.Bold);
            section.ForeColor = ApplicationTheme.PrimaryText;
            strip.Controls.Add(section);

            Button workspace = new Button();
            workspace.Left = 96;
            workspace.Top = 3;
            workspace.Width = 112;
            workspace.Height = 26;
            workspace.TabStop = true;
            ApplicationTheme.ApplyButton(workspace);
            workspace.Click += delegate { ShowOfficeWorkspace(); };
            strip.Controls.Add(workspace);

            Label formats = new Label();
            formats.Left = 220;
            formats.Top = 6;
            formats.Width = 520;
            formats.Height = 20;
            formats.ForeColor = ApplicationTheme.SecondaryText;
            strip.Controls.Add(formats);

            Label shortcut = new Label();
            shortcut.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            shortcut.Top = 6;
            shortcut.Width = 210;
            shortcut.Height = 20;
            shortcut.Left = Math.Max(746, toolbar.ClientSize.Width - shortcut.Width - 12);
            shortcut.TextAlign = ContentAlignment.MiddleRight;
            shortcut.ForeColor = ApplicationTheme.SecondaryText;
            strip.Controls.Add(shortcut);

            EventHandler applyLanguage = delegate
            {
                bool korean = UiLocalization.CurrentLanguage == AppLanguage.Korean;
                section.Text = korean ? "문서" : "Documents";
                workspace.Text = korean ? "새로 만들기 / 열기" : "New / Open";
                formats.Text = korean
                    ? "PPTX  ·  DOCX  ·  XLSX  ·  HWPX 실험적  ·  HWP 읽기 전용"
                    : "PPTX  ·  DOCX  ·  XLSX  ·  HWPX experimental  ·  HWP read-only";
                shortcut.Text = korean
                    ? "작업 공간  Ctrl+Alt+N"
                    : "Workspace  Ctrl+Alt+N";
            };

            applyLanguage(null, EventArgs.Empty);
            UiLocalization.LanguageChanged += applyLanguage;

            FormClosed += delegate
            {
                UiLocalization.LanguageChanged -= applyLanguage;
            };

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
