using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace PptxViewer
{
    internal static class AdvancedEditorFidelityExtension
    {
        private sealed class PreviewState
        {
            public AdvancedPresentationEditorForm Editor;
            public PresentationEditSession Session;
            public AdvancedPresentationCanvas Canvas;
            public ListBox SlideList;
            public Button ToggleButton;
            public Panel Overlay;
            public PictureBox Picture;
            public Label Banner;
            public Button RefreshButton;
            public readonly List<string> RenderedSlides = new List<string>();
            public string TempDirectory;
            public bool PreviewMode;
        }

        private static readonly Dictionary<AdvancedPresentationEditorForm, PreviewState> States =
            new Dictionary<AdvancedPresentationEditorForm, PreviewState>();
        private static bool initialized;

        public static void Initialize()
        {
            if (initialized)
                return;

            initialized = true;
            Application.Idle += OnApplicationIdle;
            UiLocalization.LanguageChanged += OnLanguageChanged;
        }

        public static void TryAttach(AdvancedPresentationEditorForm editor)
        {
            if (editor == null || editor.IsDisposed || States.ContainsKey(editor))
                return;

            PresentationEditSession session = GetSession(editor);
            AdvancedPresentationCanvas canvas = FindControl<AdvancedPresentationCanvas>(editor);
            ListBox slideList = FindControl<ListBox>(editor);
            Panel toolbar = FindMainToolbar(editor);
            Panel extensionHost =
                FindNamedPanel(
                    editor,
                    "AdvancedEditorToolbarExtensionHost");

            if (session == null || canvas == null || slideList == null || toolbar == null)
                return;

            PreviewState state = new PreviewState();
            state.Editor = editor;
            state.Session = session;
            state.Canvas = canvas;
            state.SlideList = slideList;

            Button toggle = new Button();
            toggle.TabStop = true;
            ApplicationTheme.ApplyButton(toggle);

            if (extensionHost != null)
            {
                toggle.Dock = DockStyle.Fill;
                toggle.Margin = Padding.Empty;
                extensionHost.Controls.Add(toggle);
            }
            else
            {
                toggle.Width = 82;
                toggle.Height = 32;
                toggle.Top = 8;
                toggle.Left =
                    Math.Max(
                        8,
                        toolbar.ClientSize.Width -
                        toggle.Width -
                        8);
                toggle.Anchor =
                    AnchorStyles.Top |
                    AnchorStyles.Right;
                toolbar.Controls.Add(toggle);
            }

            state.ToggleButton = toggle;

            CreateOverlay(state);
            UpdateLanguage(state);

            toggle.Click += delegate
            {
                if (state.PreviewMode)
                    SetPreviewMode(state, false);
                else
                    BuildAndShowPreview(state);
            };

            state.RefreshButton.Click += delegate
            {
                BuildAndShowPreview(state);
            };

            slideList.SelectedIndexChanged += delegate
            {
                if (state.PreviewMode)
                    ShowSelectedSlide(state);
            };

            editor.FormClosed += delegate
            {
                Cleanup(state);
                States.Remove(editor);
            };

            States[editor] = state;
        }

        private static void OnApplicationIdle(object sender, EventArgs e)
        {
            List<AdvancedPresentationEditorForm> editors =
                new List<AdvancedPresentationEditorForm>();

            for (int i = 0; i < Application.OpenForms.Count; i++)
            {
                AdvancedPresentationEditorForm editor =
                    Application.OpenForms[i] as AdvancedPresentationEditorForm;
                if (editor != null)
                    editors.Add(editor);
            }

            for (int i = 0; i < editors.Count; i++)
                TryAttach(editors[i]);
        }

        private static void OnLanguageChanged(object sender, EventArgs e)
        {
            List<PreviewState> states = new List<PreviewState>();
            foreach (KeyValuePair<AdvancedPresentationEditorForm, PreviewState> item in States)
            {
                if (item.Value != null &&
                    item.Value.Editor != null &&
                    !item.Value.Editor.IsDisposed)
                {
                    states.Add(item.Value);
                }
            }

            for (int i = 0; i < states.Count; i++)
                UpdateLanguage(states[i]);
        }

        private static PresentationEditSession GetSession(
            AdvancedPresentationEditorForm editor)
        {
            try
            {
                FieldInfo field = typeof(AdvancedPresentationEditorForm).GetField(
                    "session",
                    BindingFlags.Instance | BindingFlags.NonPublic);
                return field == null
                    ? null
                    : field.GetValue(editor) as PresentationEditSession;
            }
            catch
            {
                return null;
            }
        }

        private static void CreateOverlay(PreviewState state)
        {
            Panel overlay = new Panel();
            overlay.Dock = DockStyle.Fill;
            overlay.BackColor = ApplicationTheme.Canvas;
            overlay.Visible = false;

            PictureBox picture = new PictureBox();
            picture.Dock = DockStyle.Fill;
            picture.SizeMode = PictureBoxSizeMode.Zoom;
            picture.BackColor = Color.FromArgb(12, 14, 17);
            overlay.Controls.Add(picture);

            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 38;
            header.BackColor = ApplicationTheme.Toolbar;
            overlay.Controls.Add(header);
            header.BringToFront();

            Label banner = new Label();
            banner.Left = 12;
            banner.Top = 9;
            banner.Height = 20;
            banner.AutoEllipsis = true;
            banner.ForeColor = ApplicationTheme.SecondaryText;
            header.Controls.Add(banner);

            Button refresh = new Button();
            refresh.Height = 28;
            refresh.Top = 5;
            refresh.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Right;
            ApplicationTheme.ApplyButton(refresh);
            header.Controls.Add(refresh);

            Action layoutHeader = delegate
            {
                int width =
                    Math.Max(
                        0,
                        header.ClientSize.Width);
                bool showRefresh =
                    width >= 280;

                refresh.Visible =
                    showRefresh;

                if (showRefresh)
                {
                    refresh.Width =
                        width < 440
                            ? 110
                            : 142;
                    refresh.Left =
                        Math.Max(
                            10,
                            width -
                            refresh.Width -
                            10);
                    banner.Width =
                        Math.Max(
                            40,
                            refresh.Left -
                            banner.Left -
                            10);
                }
                else
                {
                    banner.Width =
                        Math.Max(
                            40,
                            width -
                            banner.Left -
                            12);
                }
            };

            header.Resize += delegate
            {
                layoutHeader();
            };
            layoutHeader();

            state.Canvas.Controls.Add(overlay);
            state.Overlay = overlay;
            state.Picture = picture;
            state.Banner = banner;
            state.RefreshButton = refresh;
        }

        private static void BuildAndShowPreview(PreviewState state)
        {
            if (state == null ||
                state.Editor == null ||
                state.Editor.IsDisposed ||
                state.Session == null ||
                state.Session.Document == null)
            {
                return;
            }

            Cursor previousCursor = state.Editor.Cursor;
            state.Editor.Cursor = Cursors.WaitCursor;

            try
            {
                DisposePreviewImage(state);
                DeleteTempDirectory(state.TempDirectory);
                state.RenderedSlides.Clear();

                string root = Path.Combine(
                    Path.GetTempPath(),
                    "PowerPointLite",
                    "EditorPreview",
                    Guid.NewGuid().ToString("N"));
                string renderDirectory = Path.Combine(root, "render");
                string packagePath = Path.Combine(root, "preview.pptx");

                Directory.CreateDirectory(renderDirectory);
                PresentationPackageWriter.Save(
                    state.Session.Document,
                    packagePath);

                List<string> slides = InternalPptxRenderer.Render(
                    packagePath,
                    renderDirectory);

                state.TempDirectory = root;
                state.RenderedSlides.AddRange(slides);

                SetPreviewMode(state, true);
                ShowSelectedSlide(state);
            }
            catch (Exception ex)
            {
                SetPreviewMode(state, false);
                MessageBox.Show(
                    state.Editor,
                    Localized(
                        "Viewer 렌더러 미리보기를 만들지 못했습니다.\r\n\r\n",
                        "Could not build the viewer renderer preview.\r\n\r\n") + ex.Message,
                    Localized("미리보기 실패", "Preview failed"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                state.Editor.Cursor = previousCursor;
            }
        }

        private static void SetPreviewMode(
            PreviewState state,
            bool enabled)
        {
            if (state == null)
                return;

            state.PreviewMode = enabled;
            state.Overlay.Visible = enabled;

            if (enabled)
                state.Overlay.BringToFront();

            UpdateLanguage(state);
        }

        private static void ShowSelectedSlide(PreviewState state)
        {
            if (state == null || !state.PreviewMode)
                return;

            int index = state.SlideList.SelectedIndex;
            if (index < 0 || index >= state.RenderedSlides.Count)
            {
                DisposePreviewImage(state);
                return;
            }

            string path = state.RenderedSlides[index];
            if (!File.Exists(path))
                return;

            try
            {
                using (Image source = Image.FromFile(path))
                {
                    Image replacement = new Bitmap(source);
                    DisposePreviewImage(state);
                    state.Picture.Image = replacement;
                }

                state.Banner.Text = Localized(
                    "Viewer 렌더러 미리보기 · 슬라이드 " + (index + 1).ToString() +
                    " · 편집 내용을 바꾼 뒤에는 새로고침하세요",
                    "Viewer renderer preview · slide " + (index + 1).ToString() +
                    " · refresh after editing changes");
                state.Overlay.BringToFront();
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine(
                    "Editor fidelity preview image failed: " + ex.Message);
            }
        }

        private static void UpdateLanguage(PreviewState state)
        {
            if (state == null)
                return;

            if (state.ToggleButton != null)
            {
                state.ToggleButton.Text = state.PreviewMode
                    ? Localized("편집", "Edit")
                    : Localized("미리보기", "Preview");
            }

            if (state.RefreshButton != null)
                state.RefreshButton.Text = Localized("미리보기 새로고침", "Refresh Preview");

            if (state.Banner != null && !state.PreviewMode)
            {
                state.Banner.Text = Localized(
                    "Viewer 렌더러와 동일한 Internal OpenXML 결과를 확인합니다",
                    "Preview the same Internal OpenXML result used by the Viewer");
            }
            else if (state.Banner != null && state.PreviewMode)
            {
                ShowSelectedSlide(state);
            }
        }

        private static string Localized(string korean, string english)
        {
            return UiLocalization.CurrentLanguage == AppLanguage.Korean
                ? korean
                : english;
        }

        private static void DisposePreviewImage(PreviewState state)
        {
            if (state == null || state.Picture == null || state.Picture.Image == null)
                return;

            Image image = state.Picture.Image;
            state.Picture.Image = null;
            image.Dispose();
        }

        private static void Cleanup(PreviewState state)
        {
            if (state == null)
                return;

            DisposePreviewImage(state);
            DeleteTempDirectory(state.TempDirectory);
            state.RenderedSlides.Clear();
            state.TempDirectory = null;
        }

        private static void DeleteTempDirectory(string path)
        {
            if (string.IsNullOrEmpty(path))
                return;

            try
            {
                if (Directory.Exists(path))
                    Directory.Delete(path, true);
            }
            catch { }
        }

        private static T FindControl<T>(Control root)
            where T : Control
        {
            if (root == null)
                return null;

            T direct = root as T;
            if (direct != null)
                return direct;

            for (int i = 0; i < root.Controls.Count; i++)
            {
                T found = FindControl<T>(root.Controls[i]);
                if (found != null)
                    return found;
            }

            return null;
        }

        private static Panel FindNamedPanel(
            Control root,
            string name)
        {
            if (root == null ||
                string.IsNullOrEmpty(name))
            {
                return null;
            }

            Panel panel =
                root as Panel;

            if (panel != null &&
                string.Equals(
                    panel.Name,
                    name,
                    StringComparison.Ordinal))
            {
                return panel;
            }

            for (int i = 0;
                 i < root.Controls.Count;
                 i++)
            {
                Panel found =
                    FindNamedPanel(
                        root.Controls[i],
                        name);

                if (found != null)
                    return found;
            }

            return null;
        }

        private static Panel FindMainToolbar(Control root)
        {
            if (root == null)
                return null;

            Panel panel = root as Panel;
            if (panel != null &&
                panel.Dock == DockStyle.Top &&
                panel.Height >= 44 &&
                panel.Height <= 56)
            {
                int buttonCount = 0;
                for (int i = 0; i < panel.Controls.Count; i++)
                {
                    if (panel.Controls[i] is Button)
                        buttonCount++;
                }

                if (buttonCount >= 5)
                    return panel;
            }

            for (int i = 0; i < root.Controls.Count; i++)
            {
                Panel found = FindMainToolbar(root.Controls[i]);
                if (found != null)
                    return found;
            }

            return null;
        }
    }
}
