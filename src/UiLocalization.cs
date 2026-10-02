using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal enum AppLanguage
    {
        Korean,
        English
    }

    internal static class UiLanguage
    {
        private static readonly Dictionary<string, string> EnglishToKorean =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> KoreanToEnglish =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<Control, bool> HookedControls =
            new Dictionary<Control, bool>();
        private static bool initialized;
        private static AppLanguage currentLanguage;

        static UiLanguage()
        {
            Add("Open", "열기");
            Add("Fit", "창 맞춤");
            Add("Full", "전체 화면");
            Add("F5 Show", "F5 발표");
            Add("TOC", "목차");
            Add("Print", "인쇄");
            Add("Recent", "최근 파일");
            Add("View", "보기");
            Add("Presenter", "발표자 보기");
            Add("Auto TOC", "자동 목차");
            Add("Auto TOC*", "자동 목차*");
            Add("Next", "다음");
            Add("Previous", "이전");
            Add("Go to slide...", "슬라이드로 이동...");
            Add("All slides...", "모든 슬라이드...");
            Add("Speaker notes", "발표자 노트");
            Add("Black screen", "검은 화면");
            Add("White screen", "흰 화면");
            Add("End show", "발표 종료");
            Add("Documents", "문서");
            Add("New / Open", "새로 만들기 / 열기");
            Add("Language", "언어");
            Add("Workspace  Ctrl+Alt+N", "문서 작업공간  Ctrl+Alt+N");
            Add("PPTX  ·  DOCX  ·  XLSX  ·  HWPX experimental  ·  HWP read-only",
                "PPTX  ·  DOCX  ·  XLSX  ·  HWPX 실험적  ·  HWP 읽기 전용");
            Add("Document Workspace", "문서 작업공간");
            Add("Create, open, export, or convert a document", "문서를 만들거나 열고, 내보내거나 변환합니다");
            Add("Presentation", "프레젠테이션");
            Add("Document", "문서");
            Add("Spreadsheet", "스프레드시트");
            Add("Korean XML document", "한글 XML 문서");
            Add("OpenDocument text", "OpenDocument 문서");
            Add("OpenDocument sheet", "OpenDocument 스프레드시트");
            Add("OpenDocument presentation", "OpenDocument 프레젠테이션");
            Add("Safety first", "문서 안전 우선");
            Add("Open existing editable file...", "기존 편집 가능 파일 열기...");
            Add("Export to PDF...", "PDF로 내보내기...");
            Add("Convert format...", "형식 변환...");
            Add("Font license inspector...", "폰트 라이선스 검사...");
            Add("Cancel", "취소");
            Add("Presentation Editor", "프레젠테이션 편집기");
            Add("Save", "저장");
            Add("Save As", "다른 이름으로 저장");
            Add("Undo", "실행 취소");
            Add("Redo", "다시 실행");
            Add("Text", "텍스트");
            Add("Image", "이미지");
            Add("Shape", "도형");
            Add("Delete", "삭제");
            Add("Copy", "복사");
            Add("Paste", "붙여넣기");
            Add("+ Slide", "+ 슬라이드");
            Add("- Slide", "- 슬라이드");
            Add("Up", "위로");
            Add("Down", "아래로");
            Add("Object properties", "개체 속성");
            Add("No object selected", "선택된 개체 없음");
            Add("Text box", "텍스트 상자");
            Add("Font", "글꼴");
            Add("Size", "크기");
            Add("Bold", "굵게");
            Add("Italic", "기울임");
            Add("Alignment", "정렬");
            Add("Left", "왼쪽");
            Add("Center", "가운데");
            Add("Right", "오른쪽");
            Add("Text color", "글자 색");
            Add("Fill", "채우기");
            Add("Line", "선");
            Add("Slides", "슬라이드");
            Add("Table", "표");
            Add("Rectangle", "사각형");
            Add("Rounded rectangle", "둥근 사각형");
            Add("Ellipse", "타원");
            Add("Triangle", "삼각형");
            Add("Diamond", "마름모");
            Add("Renderer Preview", "렌더러 미리보기");
            Add("Edit View", "편집 보기");
            Add("Refresh Preview", "미리보기 새로고침");
            Add("Font License Inspector", "폰트 라이선스 검사");
            Add("OpenType / TrueType embedding metadata", "OpenType / TrueType 포함 권한 메타데이터");
            Add("Browse...", "찾아보기...");
            Add("Language Settings", "언어 설정");
            Add("Interface language", "인터페이스 언어");
            Add("Korean", "한국어");
            Add("English", "English");
            Add("Apply", "적용");
            Add("Close", "닫기");
            Add("Unsaved", "저장 안 됨");
            Add("Saved", "저장됨");
            Add("Modified", "수정됨");
            Add("Save failed", "저장 실패");
            Add("Unsaved presentation", "저장되지 않은 프레젠테이션");
            Add("Save changes before closing?", "닫기 전에 변경 내용을 저장할까요?");
            Add("Open document", "문서 열기");
            Add("Document could not be opened safely", "문서를 안전하게 열 수 없음");
            Add("Edit presentation", "프레젠테이션 편집");
            Add("Editing is not safe yet", "아직 안전하게 편집할 수 없음");
            Add("PPTX compatibility", "PPTX 호환성");
            Add("Compatibility scan failed", "호환성 검사 실패");
            Add("Presentation Info", "프레젠테이션 정보");
            Add("Speaker Notes", "발표자 노트");
            Add("Presenter View", "발표자 보기");
            Add("Auto-hide slide thumbnails", "슬라이드 썸네일 자동 숨김");
            Add("Slide Sorter", "슬라이드 정렬 보기");
            Add("Find...", "찾기...");
            Add("Fit to window", "창에 맞춤");
            Add("Full screen", "전체 화면");
            Add("Presentation info", "프레젠테이션 정보");
            Add("Pointer Options", "포인터 옵션");
            Add("Arrow", "화살표");
            Add("Laser Pointer", "레이저 포인터");
            Add("Pen", "펜");
            Add("Highlighter", "형광펜");
            Add("Eraser", "지우개");
            Add("Erase Ink on Slide", "현재 슬라이드 필기 지우기");
            Add("Erase All Ink", "모든 필기 지우기");

            currentLanguage = LoadLanguage();
        }

        public static AppLanguage Current
        {
            get { return currentLanguage; }
        }

        public static bool IsKorean
        {
            get { return currentLanguage == AppLanguage.Korean; }
        }

        public static string T(string korean, string english)
        {
            return IsKorean ? korean : english;
        }

        public static void ApplyTree(Control root)
        {
            EnsureInitialized();
            if (root == null)
                return;

            ApplyControl(root);
            HookControl(root);
        }

        public static void ApplyAllOpenForms()
        {
            EnsureInitialized();

            for (int i = 0; i < Application.OpenForms.Count; i++)
                ApplyTree(Application.OpenForms[i]);
        }

        public static void ShowSettings(IWin32Window owner)
        {
            using (Form form = new Form())
            {
                form.Text = T("언어 설정", "Language Settings");
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowInTaskbar = false;
                form.Width = 430;
                form.Height = 220;
                form.BackColor = ApplicationTheme.Window;
                form.ForeColor = ApplicationTheme.PrimaryText;
                form.Font = new System.Drawing.Font("Segoe UI", 9f);

                Label label = new Label();
                label.Left = 24;
                label.Top = 24;
                label.Width = 350;
                label.Height = 24;
                label.Text = T("인터페이스 언어", "Interface language");
                label.Font = new System.Drawing.Font(form.Font, System.Drawing.FontStyle.Bold);
                form.Controls.Add(label);

                ComboBox combo = new ComboBox();
                combo.Left = 24;
                combo.Top = 56;
                combo.Width = 360;
                combo.DropDownStyle = ComboBoxStyle.DropDownList;
                combo.Items.Add("한국어");
                combo.Items.Add("English");
                combo.SelectedIndex = currentLanguage == AppLanguage.Korean ? 0 : 1;
                form.Controls.Add(combo);

                Label hint = new Label();
                hint.Left = 24;
                hint.Top = 92;
                hint.Width = 360;
                hint.Height = 34;
                hint.Text = T(
                    "기본 언어는 한국어입니다. 설정은 이 PC에 저장됩니다.",
                    "Korean is the default. The setting is stored on this PC.");
                hint.ForeColor = ApplicationTheme.SecondaryText;
                form.Controls.Add(hint);

                Button apply = new Button();
                apply.Left = 204;
                apply.Top = 132;
                apply.Width = 86;
                apply.Height = 30;
                apply.Text = T("적용", "Apply");
                ApplicationTheme.ApplyButton(apply);
                form.Controls.Add(apply);

                Button close = new Button();
                close.Left = 298;
                close.Top = 132;
                close.Width = 86;
                close.Height = 30;
                close.Text = T("닫기", "Close");
                ApplicationTheme.ApplyButton(close);
                close.Click += delegate { form.Close(); };
                form.Controls.Add(close);

                apply.Click += delegate
                {
                    currentLanguage = combo.SelectedIndex == 1
                        ? AppLanguage.English
                        : AppLanguage.Korean;
                    SaveLanguage();
                    ApplyAllOpenForms();
                    form.Text = T("언어 설정", "Language Settings");
                    label.Text = T("인터페이스 언어", "Interface language");
                    hint.Text = T(
                        "기본 언어는 한국어입니다. 설정은 이 PC에 저장됩니다.",
                        "Korean is the default. The setting is stored on this PC.");
                    apply.Text = T("적용", "Apply");
                    close.Text = T("닫기", "Close");
                };

                form.AcceptButton = apply;
                form.CancelButton = close;
                form.ShowDialog(owner);
            }
        }

        private static void EnsureInitialized()
        {
            if (initialized)
                return;

            initialized = true;
            Application.Idle += OnApplicationIdle;
        }

        private static void OnApplicationIdle(object sender, EventArgs e)
        {
            for (int i = 0; i < Application.OpenForms.Count; i++)
            {
                Form form = Application.OpenForms[i];
                if (!HookedControls.ContainsKey(form))
                    ApplyTree(form);

                AdvancedPresentationEditorForm editor =
                    form as AdvancedPresentationEditorForm;
                if (editor != null)
                    AdvancedEditorFidelityExtension.TryAttach(editor);
            }
        }

        private static void HookControl(Control control)
        {
            if (control == null || HookedControls.ContainsKey(control))
                return;

            HookedControls[control] = true;
            control.ControlAdded += delegate(object sender, ControlEventArgs e)
            {
                ApplyTree(e.Control);
            };
            control.Disposed += delegate
            {
                HookedControls.Remove(control);
            };

            for (int i = 0; i < control.Controls.Count; i++)
                HookControl(control.Controls[i]);
        }

        private static void ApplyControl(Control control)
        {
            if (control == null)
                return;

            if (!string.IsNullOrEmpty(control.Text))
                control.Text = Translate(control.Text);

            ComboBox combo = control as ComboBox;
            if (combo != null)
            {
                int selected = combo.SelectedIndex;
                for (int i = 0; i < combo.Items.Count; i++)
                {
                    string item = combo.Items[i] as string;
                    if (!string.IsNullOrEmpty(item))
                        combo.Items[i] = Translate(item);
                }
                if (selected >= 0 && selected < combo.Items.Count)
                    combo.SelectedIndex = selected;
            }

            ContextMenuStrip menu = control as ContextMenuStrip;
            if (menu != null)
                ApplyToolStripItems(menu.Items);

            for (int i = 0; i < control.Controls.Count; i++)
                ApplyControl(control.Controls[i]);
        }

        private static void ApplyToolStripItems(ToolStripItemCollection items)
        {
            if (items == null)
                return;

            for (int i = 0; i < items.Count; i++)
            {
                ToolStripItem item = items[i];
                if (item == null)
                    continue;

                if (!string.IsNullOrEmpty(item.Text))
                    item.Text = Translate(item.Text);

                ToolStripDropDownItem dropDown = item as ToolStripDropDownItem;
                if (dropDown != null)
                    ApplyToolStripItems(dropDown.DropDownItems);
            }
        }

        private static string Translate(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            string translated;
            if (currentLanguage == AppLanguage.Korean)
            {
                if (EnglishToKorean.TryGetValue(text, out translated))
                    return translated;

                if (text.StartsWith("Engine: ", StringComparison.OrdinalIgnoreCase))
                    return "엔진: " + text.Substring(8);
                if (text.EndsWith(" - Editor", StringComparison.OrdinalIgnoreCase))
                    return text.Substring(0, text.Length - 9) + " - 편집기";
                if (text.StartsWith("Slide ", StringComparison.OrdinalIgnoreCase))
                    return "슬라이드 " + text.Substring(6);
                if (text.StartsWith("Modified  •  ", StringComparison.OrdinalIgnoreCase))
                    return "수정됨  •  " + text.Substring(13);
                if (text.StartsWith("Saved  •  ", StringComparison.OrdinalIgnoreCase))
                    return "저장됨  •  " + text.Substring(10);
            }
            else
            {
                if (KoreanToEnglish.TryGetValue(text, out translated))
                    return translated;

                if (text.StartsWith("엔진: ", StringComparison.OrdinalIgnoreCase))
                    return "Engine: " + text.Substring(4);
                if (text.EndsWith(" - 편집기", StringComparison.OrdinalIgnoreCase))
                    return text.Substring(0, text.Length - 6) + " - Editor";
                if (text.StartsWith("슬라이드 ", StringComparison.OrdinalIgnoreCase))
                    return "Slide " + text.Substring(5);
                if (text.StartsWith("수정됨  •  ", StringComparison.OrdinalIgnoreCase))
                    return "Modified  •  " + text.Substring(9);
                if (text.StartsWith("저장됨  •  ", StringComparison.OrdinalIgnoreCase))
                    return "Saved  •  " + text.Substring(9);
            }

            return text;
        }

        private static void Add(string english, string korean)
        {
            EnglishToKorean[english] = korean;
            KoreanToEnglish[korean] = english;
        }

        private static AppLanguage LoadLanguage()
        {
            try
            {
                string path = SettingsPath;
                if (!File.Exists(path))
                    return AppLanguage.Korean;

                string[] lines = File.ReadAllLines(path);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (string.Equals(line, "language=en", StringComparison.OrdinalIgnoreCase))
                        return AppLanguage.English;
                    if (string.Equals(line, "language=ko", StringComparison.OrdinalIgnoreCase))
                        return AppLanguage.Korean;
                }
            }
            catch { }

            return AppLanguage.Korean;
        }

        private static void SaveLanguage()
        {
            try
            {
                string path = SettingsPath;
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(
                    path,
                    currentLanguage == AppLanguage.English
                        ? "language=en" + Environment.NewLine
                        : "language=ko" + Environment.NewLine);
            }
            catch { }
        }

        private static string SettingsPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointLite",
                    "settings.ini");
            }
        }
    }
}
