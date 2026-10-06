using System;
using System.Drawing;
using System.Windows.Forms;

namespace PptxViewer
{
    public sealed partial class MainForm : Form
    {
        private Button languageButton;
        private ContextMenuStrip languageMenu;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            UiLocalization.Initialize();
            CreateLanguageSelector();
            InitializePresentationPointerTools();
            InitializeSlideShowTiming();

            UiLocalization.LanguageChanged += OnUiLanguageChanged;
            FormClosed += delegate
            {
                UiLocalization.LanguageChanged -= OnUiLanguageChanged;
            };

            ApplyViewerLanguage();

            // MainFormWorkspaceToolbar performs its own OnShown layout after
            // base.OnShown. Defer once so the language selector remains clear
            // of the engine label after that layout has completed.
            Shown += delegate
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    LayoutViewerToolbar();
                    ApplyViewerLanguage();
                });
            };

            toolbar.Resize += delegate
            {
                LayoutViewerToolbar();
            };
        }

        private void CreateLanguageSelector()
        {
            if (languageButton != null)
                return;

            languageButton = new Button();
            languageButton.Top = 8;
            languageButton.Width = 82;
            languageButton.Height = 30;
            languageButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            languageButton.TextAlign = ContentAlignment.MiddleCenter;
            languageButton.UseCompatibleTextRendering = true;
            languageButton.Padding = new Padding(0, 1, 0, 0);
            languageButton.AutoSize = false;
            languageButton.AutoEllipsis = true;
            languageButton.Tag = "Language";
            ApplicationTheme.ApplyButton(languageButton);

            languageMenu = new ContextMenuStrip();

            ToolStripMenuItem korean = new ToolStripMenuItem("한국어");
            ToolStripMenuItem english = new ToolStripMenuItem("English");

            korean.Click += delegate
            {
                UiLocalization.SetLanguage(AppLanguage.Korean);
            };

            english.Click += delegate
            {
                UiLocalization.SetLanguage(AppLanguage.English);
            };

            languageMenu.Items.Add(korean);
            languageMenu.Items.Add(english);

            languageButton.Click += delegate
            {
                RefreshLanguageMenuChecks();
                languageMenu.Show(
                    languageButton,
                    new Point(0, languageButton.Height));
            };

            toolbar.Controls.Add(languageButton);
            languageButton.BringToFront();
            LayoutViewerToolbar();
        }

        private void LayoutViewerToolbar()
        {
            if (toolbar == null)
                return;

            int buttonHeight = 30;
            Control documentStrip = toolbar.Controls["DocumentWorkspaceStrip"];
            int reservedBottom = documentStrip != null && documentStrip.Visible
                ? documentStrip.Height
                : 0;
            int topBandHeight = Math.Max(
                buttonHeight,
                toolbar.ClientSize.Height - reservedBottom);
            int top = Math.Max(0, (topBandHeight - buttonHeight) / 2);
            int x = 8;
            int gap = 6;

            string[] order = new string[]
            {
                "Open",
                "<",
                ">",
                "Fit",
                "100%",
                "Full",
                "F5 Show",
                "TOC",
                "Print",
                "Recent",
                "View",
                "Presenter",
                "Workspace",
                "Auto TOC"
            };

            bool hasWorkspace = FindToolbarButton("Workspace") != null;

            for (int i = 0; i < order.Length; i++)
            {
                Button button = FindToolbarButton(order[i]);
                if (button == null)
                    continue;

                if (order[i] == "Auto TOC" && hasWorkspace)
                {
                    button.Visible = false;
                    continue;
                }

                button.Visible = true;
                NormalizeToolbarButton(button);

                int minimumWidth = GetToolbarButtonMinimumWidth(order[i]);
                Size measured = TextRenderer.MeasureText(
                    button.Text ?? "",
                    button.Font,
                    new Size(500, buttonHeight),
                    TextFormatFlags.SingleLine |
                    TextFormatFlags.NoPadding);

                int desiredWidth = Math.Max(
                    minimumWidth,
                    Math.Min(118, measured.Width + 20));

                button.SetBounds(
                    x,
                    top,
                    desiredWidth,
                    buttonHeight);

                x += desiredWidth + gap;
            }

            if (zoomTrack != null)
            {
                zoomTrack.Left = x + 8;
                zoomTrack.Top = Math.Max(
                    0,
                    (topBandHeight - zoomTrack.Height) / 2);
                x = zoomTrack.Right + 14;
            }

            if (languageButton != null)
            {
                NormalizeToolbarButton(languageButton);
                languageButton.Width = 82;
                languageButton.Height = buttonHeight;
                languageButton.Top = top;
                languageButton.Left = Math.Max(
                    8,
                    toolbar.ClientSize.Width - languageButton.Width - 8);
                languageButton.BringToFront();
            }

            if (engineLabel != null)
            {
                int rightLimit = languageButton != null
                    ? languageButton.Left - 12
                    : toolbar.ClientSize.Width - 8;

                int available = Math.Max(0, rightLimit - x);

                engineLabel.Left = x;
                engineLabel.Top = Math.Max(
                    0,
                    (topBandHeight - 22) / 2);
                engineLabel.Height = 22;
                engineLabel.Width = Math.Max(0, available);
                engineLabel.Visible = available >= 70;
            }
        }

        private Button FindToolbarButton(string key)
        {
            if (toolbar == null || string.IsNullOrEmpty(key))
                return null;

            for (int i = 0; i < toolbar.Controls.Count; i++)
            {
                Button button = toolbar.Controls[i] as Button;
                if (button == null || object.ReferenceEquals(button, languageButton))
                    continue;

                string tag = button.Tag as string;
                if (string.Equals(tag, key, StringComparison.Ordinal))
                    return button;
            }

            return null;
        }

        private static void NormalizeToolbarButton(Button button)
        {
            if (button == null)
                return;

            button.TextAlign = ContentAlignment.MiddleCenter;
            button.UseCompatibleTextRendering = true;
            button.Padding = new Padding(0, 1, 0, 0);
            button.AutoSize = false;
            button.AutoEllipsis = true;
        }

        private static int GetToolbarButtonMinimumWidth(string key)
        {
            if (key == "<" || key == ">") return 44;
            if (key == "Open") return 72;
            if (key == "F5 Show") return 90;
            if (key == "Print") return 68;
            if (key == "Recent") return 72;
            if (key == "Presenter") return 82;
            if (key == "Workspace") return 84;
            if (key == "Auto TOC") return 82;
            return 62;
        }

        private void OnUiLanguageChanged(object sender, EventArgs e)
        {
            ApplyViewerLanguage();
        }

        private void ApplyViewerLanguage()
        {
            UiLocalization.ApplyToForm(this);
            UiLocalization.ApplyToolStrip(slideshowMenu);

            if (languageMenu != null)
                UiLocalization.ApplyToolStrip(languageMenu);

            if (languageButton != null)
            {
                languageButton.Text =
                    UiLocalization.CurrentLanguage == AppLanguage.Korean
                        ? "한국어"
                        : "English";
            }

            RefreshLanguageMenuChecks();
            DetectBackends();

            // DetectBackends writes canonical English engine labels so apply
            // the selected language once more afterwards.
            if (engineLabel != null)
                engineLabel.Text = UiLocalization.Text(engineLabel.Text);

            LayoutViewerToolbar();
        }

        private void RefreshLanguageMenuChecks()
        {
            if (languageMenu == null || languageMenu.Items.Count < 2)
                return;

            ToolStripMenuItem korean = languageMenu.Items[0] as ToolStripMenuItem;
            ToolStripMenuItem english = languageMenu.Items[1] as ToolStripMenuItem;

            bool isKorean = UiLocalization.CurrentLanguage == AppLanguage.Korean;

            if (korean != null)
            {
                korean.Text = isKorean ? "한국어" : "Korean";
                korean.Checked = isKorean;
            }

            if (english != null)
            {
                english.Text = "English";
                english.Checked = !isKorean;
            }
        }
    }
}
