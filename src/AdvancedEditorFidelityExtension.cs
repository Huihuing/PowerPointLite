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

        public static void TryAttach(AdvancedPresentationEditorForm editor)
        {
            if (editor == null || editor.IsDisposed || States.ContainsKey(editor))
                return;

            PresentationEditSession session = GetSession(editor);
            AdvancedPresentationCanvas canvas = FindControl<AdvancedPresentationCanvas>(editor);
            ListBox slideList = FindControl<ListBox>(editor);
            Panel toolbar = FindMainToolbar(editor);

            if (session == null || canvas == null || slideList == null || toolbar == null)
                return;

            PreviewState state = new PreviewState();
            state.Editor = editor;
            state.Session = session;
            state.Canvas = canvas;
            state.SlideList = slideList;

            Button toggle = new Button();
            toggle.Left = 976;
            toggle.Top = 8;
            toggle.Width = 150;
            toggle.Height = 32;
            toggle.Text = "Renderer Preview";
            toggle.TabStop = true;
            ApplicationTheme.ApplyButton(toggle);
            toolbar.Controls.Add(toggle);
            state.ToggleButton = toggle;

            CreateOverlay(state);

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
            UiLanguage.ApplyTree(toggle);
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
            banner.Width = 620;
            banner.Height = 20;
            banner.ForeColor = ApplicationTheme.SecondaryText;
            banner.Text = "Viewer renderer preview · Internal OpenXML · read-only preview";
            header.Controls.Add(banner);

            Button refresh = new Button();
            refresh.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            refresh.Width = 142;
            refresh.Height = 28;
            refresh.Top = 5;
            refresh.Left = Math.Max(640, header.ClientSize.Width - refresh.Width - 10);
            refresh.Text = "Refresh Preview";
            ApplicationTheme.ApplyButton(refresh);
            header.Controls.Add(refresh);

            header.Resize += delegate
            {
                refresh.Left = Math.Max(
                    640,
                    header.ClientSize.Width - refresh.Width - 10);
            };

            state.Canvas.Controls.Add(overlay);
            state.Overlay = overlay;
            state.Picture = picture;
            state.Banner = banner;
            state.RefreshButton = refresh;

            UiLanguage.ApplyTree(overlay);
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
                    UiLanguage.T(
                        "Viewer 렌더러 미리보기를 만들지 못했습니다.\r\n\r\n",
                        "Could not build the viewer renderer preview.\r\n\r\n") + ex.Message,
                    UiLanguage.T("미리보기 실패", "Preview failed"),
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
            {
                state.Overlay.BringToFront();
                state.ToggleButton.Text = "Edit View";
                state.Banner.Text = UiLanguage.T(
                    "Viewer 렌더러 미리보기 · Internal OpenXML · 읽기 전용 미리보기",
                    "Viewer renderer preview · Internal OpenXML · read-only preview");
            }
            else
            {
                state.ToggleButton.Text = "Renderer Preview";
            }

            UiLanguage.ApplyTree(state.ToggleButton);
            UiLanguage.ApplyTree(state.Overlay);
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

                state.Banner.Text = UiLanguage.T(
                    "Viewer 렌더러 미리보기 · 슬라이드 " + (index + 1).ToString() +
                    " · 편집 내용 변경 후에는 새로고침하세요",
                    "Viewer renderer preview · slide " + (index + 1).ToString() +
                    " · refresh after editing changes");
                UiLanguage.ApplyTree(state.Banner);
                state.Overlay.BringToFront();
            }
            catch (Exception ex)
            {
                CrashReporter.WriteLine(
                    "Editor fidelity preview image failed: " + ex.Message);
            }
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
