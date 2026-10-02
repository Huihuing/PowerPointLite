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
                    LayoutLanguageSelector();
                    ApplyViewerLanguage();
                });
            };

            toolbar.Resize += delegate
            {
                LayoutLanguageSelector();
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
            LayoutLanguageSelector();
        }

        private void LayoutLanguageSelector()
        {
            if (languageButton == null || toolbar == null)
                return;

            int toolbarWidth = Math.Max(900, toolbar.ClientSize.Width);
            languageButton.Left = Math.Max(8, toolbarWidth - languageButton.Width - 8);

            // Keep the renderer label readable instead of placing controls on
            // top of each other at the default 1400 px window width.
            int engineRightLimit = languageButton.Left - 178;
            if (engineLabel != null && engineLabel.Left > engineRightLimit)
                engineLabel.Left = Math.Max(8, engineRightLimit);
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

            LayoutLanguageSelector();
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
