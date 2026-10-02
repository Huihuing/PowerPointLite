using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace PptxViewer
{
    internal enum AppLanguage
    {
        Korean,
        English
    }

    internal static class UiLocalization
    {
        private static readonly Dictionary<string, string> EnglishToKorean =
            new Dictionary<string, string>(StringComparer.Ordinal);

        private static readonly Dictionary<string, string> KoreanToEnglish =
            new Dictionary<string, string>(StringComparer.Ordinal);

        private static readonly Dictionary<Form, int> AppliedForms =
            new Dictionary<Form, int>();

        private static bool initialized;
        private static int languageGeneration = 1;
        private static AppLanguage currentLanguage = AppLanguage.Korean;

        public static event EventHandler LanguageChanged;

        static UiLocalization()
        {
            Add("Language", "언어");
            Add("Korean", "한국어");
            Add("English", "English");
            Add("Workspace", "작업 공간");

            Add("Open", "열기");
            Add("Open...", "열기...");
            Add("Save", "저장");
            Add("Save As", "다른 이름으로 저장");
            Add("Save As...", "다른 이름으로 저장...");
            Add("Cancel", "취소");
            Add("Close", "닫기");
            Add("Delete", "삭제");
            Add("Copy", "복사");
            Add("Paste", "붙여넣기");
            Add("Undo", "실행 취소");
            Add("Redo", "다시 실행");
            Add("Previous", "이전");
            Add("Next", "다음");
            Add("Up", "위로");
            Add("Down", "아래로");
            Add("Print", "인쇄");
            Add("View", "보기");
            Add("Recent", "최근 파일");
            Add("Fit", "맞춤");
            Add("Fit to window", "창에 맞춤");
            Add("Full", "전체 화면");
            Add("Full screen", "전체 화면");
            Add("F5 Show", "F5 슬라이드 쇼");
            Add("TOC", "목록");
            Add("Auto TOC", "자동 목록");
            Add("Auto TOC*", "자동 목록*");
            Add("Presenter", "발표자 보기");
            Add("Presenter View", "발표자 보기");
            Add("Speaker Notes", "발표자 노트");
            Add("Speaker notes", "발표자 노트");
            Add("Slide Sorter", "슬라이드 정렬");
            Add("Find...", "찾기...");
            Add("Go to slide...", "슬라이드로 이동...");
            Add("All slides...", "모든 슬라이드...");
            Add("Presentation info", "프레젠테이션 정보");
            Add("Black screen", "검은 화면");
            Add("White screen", "흰 화면");
            Add("End show", "슬라이드 쇼 끝내기");
            Add("Auto-hide slide thumbnails", "슬라이드 미리보기 자동 숨김");
            Add("(No recent files)", "(최근 파일 없음)");
            Add("Clear recent files", "최근 파일 지우기");
            Add("Drop a PPTX file here or click Open", "PPTX 파일을 끌어 놓거나 '열기'를 누르세요");
            Add("Engine: detecting...", "렌더링 엔진: 확인 중...");
            Add("Engine: PowerPoint Native", "렌더링 엔진: PowerPoint Native");
            Add("Engine: Internal OpenXML", "렌더링 엔진: 내부 OpenXML");
            Add("Engine: Internal OpenXML + LibreOffice fallback", "렌더링 엔진: 내부 OpenXML + LibreOffice 보조");

            Add("Document Workspace", "문서 작업 공간");
            Add("Create, open, export, or convert a document", "문서를 만들고, 열고, 내보내거나 변환합니다");
            Add("Independent readers and writers with project-owned UI. Microsoft/Hancom application assets are not bundled.",
                "프로젝트 자체 UI와 독립 Reader/Writer를 사용합니다. Microsoft/한컴 프로그램 자산은 포함하지 않습니다.");
            Add("Presentation", "프레젠테이션");
            Add("Document", "문서");
            Add("Spreadsheet", "스프레드시트");
            Add("Korean XML document", "한글 XML 문서");
            Add("OpenDocument text", "OpenDocument 텍스트");
            Add("OpenDocument sheet", "OpenDocument 스프레드시트");
            Add("OpenDocument presentation", "OpenDocument 프레젠테이션");
            Add("Slides, text, images, shapes, tables, viewer and presentation mode.",
                "슬라이드, 텍스트, 이미지, 도형, 표, 보기 및 발표 모드를 지원합니다.");
            Add("Paragraphs and styled text using the shared document model.",
                "공통 문서 모델을 사용해 문단과 서식 있는 텍스트를 편집합니다.");
            Add("Worksheets, cells and formula storage. Calculation engine is still limited.",
                "워크시트, 셀, 수식 저장을 지원합니다. 수식 계산 엔진은 아직 제한적입니다.");
            Add("Public-format reader/writer. Real Hancom compatibility still requires verification.",
                "공개 형식 기반 Reader/Writer입니다. 실제 한컴 호환성 검증이 더 필요합니다.");
            Add("ODF text reader/writer connected to the shared document editor.",
                "공통 문서 편집기에 연결된 ODF 텍스트 Reader/Writer입니다.");
            Add("ODF spreadsheet reader/writer with a lightweight worksheet editor.",
                "경량 워크시트 편집기를 포함한 ODF 스프레드시트 Reader/Writer입니다.");
            Add("ODF presentation reader/writer with independent slide editing tools.",
                "독립 슬라이드 편집 도구를 포함한 ODF 프레젠테이션 Reader/Writer입니다.");
            Add("Safety first", "문서 손실 방지 우선");
            Add("Existing files are edited or converted only when the current parser can represent them without known silent data loss.",
                "현재 파서가 알려진 데이터 손실 없이 표현할 수 있는 파일만 편집하거나 변환합니다.");
            Add("Open existing editable file...", "편집 가능한 기존 파일 열기...");
            Add("Export to PDF...", "PDF로 내보내기...");
            Add("Convert format...", "형식 변환...");
            Add("Font license inspector...", "글꼴 라이선스 검사...");
            Add("Metadata is advisory; actual font license text takes priority.",
                "글꼴 메타데이터는 참고용이며 실제 라이선스 원문이 우선합니다.");

            Add("Presentation Editor", "프레젠테이션 편집기");
            Add("Document Editor", "문서 편집기");
            Add("Spreadsheet Editor", "스프레드시트 편집기");
            Add("Object properties", "개체 속성");
            Add("No object selected", "선택한 개체 없음");
            Add("Text", "텍스트");
            Add("Image", "이미지");
            Add("Shape", "도형");
            Add("Font", "글꼴");
            Add("Size", "크기");
            Add("Bold", "굵게");
            Add("Italic", "기울임");
            Add("Underline", "밑줄");
            Add("Alignment", "정렬");
            Add("Left", "왼쪽");
            Add("Center", "가운데");
            Add("Right", "오른쪽");
            Add("Text color", "글자색");
            Add("Fill color", "채우기 색");
            Add("Line color", "선 색");
            Add("Rectangle", "사각형");
            Add("Rounded rectangle", "둥근 사각형");
            Add("Ellipse", "타원");
            Add("Triangle", "삼각형");
            Add("Diamond", "마름모");
            Add("+ Slide", "+ 슬라이드");
            Add("- Slide", "- 슬라이드");
            Add("Add", "추가");
            Add("Add slide", "슬라이드 추가");
            Add("Delete slide", "슬라이드 삭제");
            Add("Formula", "수식");
            Add("Worksheet", "워크시트");
            Add("Sheet", "시트");
            Add("New sheet", "새 시트");
            Add("Rename", "이름 바꾸기");

            Add("Current slide", "현재 슬라이드");
            Add("Next slide", "다음 슬라이드");
            Add("Notes", "노트");
            Add("Elapsed", "경과 시간");
            Add("Start Show", "슬라이드 쇼 시작");
            Add("Go To", "이동");

            Add("Full Slide", "전체 슬라이드");
            Add("Notes Page", "노트 페이지");
            Add("Handout 2", "유인물 2슬라이드");
            Add("Handout 4", "유인물 4슬라이드");
            Add("Handout 6", "유인물 6슬라이드");

            Add("Font License Inspector", "글꼴 라이선스 검사");
            Add("Select font file...", "글꼴 파일 선택...");
            Add("Embedding metadata", "임베딩 메타데이터");
            Add("Actual license text takes priority over font metadata.",
                "글꼴 메타데이터보다 실제 라이선스 원문이 우선합니다.");

            Add("Open document", "문서 열기");
            Add("Open failed", "열기 실패");
            Add("Find", "찾기");
            Add("Go to slide", "슬라이드로 이동");
            Add("Print failed", "인쇄 실패");
            Add("Editing is not safe yet", "아직 안전하게 편집할 수 없음");
            Add("Compatibility scan failed", "호환성 검사 실패");
            Add("PPTX compatibility", "PPTX 호환성");
        }

        public static AppLanguage CurrentLanguage
        {
            get
            {
                EnsureInitialized();
                return currentLanguage;
            }
        }

        public static bool IsKorean
        {
            get { return CurrentLanguage == AppLanguage.Korean; }
        }

        public static void Initialize()
        {
            if (initialized)
                return;

            initialized = true;
            currentLanguage = LoadLanguage();
            Application.Idle += OnApplicationIdle;
        }

        public static string Text(string value)
        {
            EnsureInitialized();

            if (string.IsNullOrEmpty(value))
                return value;

            string translated;

            if (currentLanguage == AppLanguage.Korean)
            {
                if (EnglishToKorean.TryGetValue(value, out translated))
                    return translated;
            }
            else
            {
                if (KoreanToEnglish.TryGetValue(value, out translated))
                    return translated;
            }

            return value;
        }

        public static void SetLanguage(AppLanguage language)
        {
            EnsureInitialized();

            if (currentLanguage != language)
            {
                currentLanguage = language;
                languageGeneration++;
                SaveLanguage(language);
            }

            ApplyOpenForms();

            EventHandler handler = LanguageChanged;
            if (handler != null)
                handler(null, EventArgs.Empty);
        }

        public static void ApplyOpenForms()
        {
            List<Form> forms = new List<Form>();

            try
            {
                for (int i = 0; i < Application.OpenForms.Count; i++)
                {
                    Form form = Application.OpenForms[i];
                    if (form != null)
                        forms.Add(form);
                }
            }
            catch
            {
            }

            for (int i = 0; i < forms.Count; i++)
            {
                Form form = forms[i];
                if (form == null || form.IsDisposed)
                    continue;

                int appliedGeneration;
                if (AppliedForms.TryGetValue(form, out appliedGeneration) &&
                    appliedGeneration == languageGeneration)
                {
                    continue;
                }

                ApplyToForm(form);
                AppliedForms[form] = languageGeneration;
            }

            List<Form> stale = new List<Form>();
            foreach (KeyValuePair<Form, int> item in AppliedForms)
            {
                if (item.Key == null || item.Key.IsDisposed)
                    stale.Add(item.Key);
            }

            for (int i = 0; i < stale.Count; i++)
                AppliedForms.Remove(stale[i]);
        }

        public static void ApplyToForm(Form form)
        {
            if (form == null || form.IsDisposed)
                return;

            TranslateControl(form);

            if (form.MainMenuStrip != null)
                ApplyToolStrip(form.MainMenuStrip);
        }

        public static void ApplyToolStrip(ToolStrip strip)
        {
            if (strip == null)
                return;

            for (int i = 0; i < strip.Items.Count; i++)
                TranslateToolStripItem(strip.Items[i]);
        }

        private static void TranslateControl(Control control)
        {
            if (control == null)
                return;

            bool translateText =
                control is Form ||
                control is Label ||
                control is Button ||
                control is CheckBox ||
                control is RadioButton ||
                control is GroupBox ||
                control is TabPage;

            if (translateText)
                control.Text = Text(control.Text);

            ComboBox combo = control as ComboBox;
            if (combo != null && combo.DropDownStyle == ComboBoxStyle.DropDownList)
                TranslateComboItems(combo);

            if (control.ContextMenuStrip != null)
                ApplyToolStrip(control.ContextMenuStrip);

            for (int i = 0; i < control.Controls.Count; i++)
                TranslateControl(control.Controls[i]);
        }

        private static void TranslateComboItems(ComboBox combo)
        {
            if (combo == null || combo.Items.Count == 0)
                return;

            int selectedIndex = combo.SelectedIndex;

            for (int i = 0; i < combo.Items.Count; i++)
            {
                string item = combo.Items[i] as string;
                if (item != null)
                    combo.Items[i] = Text(item);
            }

            if (selectedIndex >= 0 && selectedIndex < combo.Items.Count)
                combo.SelectedIndex = selectedIndex;
        }

        private static void TranslateToolStripItem(ToolStripItem item)
        {
            if (item == null)
                return;

            item.Text = Text(item.Text);

            ToolStripDropDownItem dropDown = item as ToolStripDropDownItem;
            if (dropDown == null)
                return;

            for (int i = 0; i < dropDown.DropDownItems.Count; i++)
                TranslateToolStripItem(dropDown.DropDownItems[i]);
        }

        private static void OnApplicationIdle(object sender, EventArgs e)
        {
            ApplyOpenForms();
        }

        private static void EnsureInitialized()
        {
            if (!initialized)
                Initialize();
        }

        private static void Add(string english, string korean)
        {
            if (!EnglishToKorean.ContainsKey(english))
                EnglishToKorean.Add(english, korean);

            if (!KoreanToEnglish.ContainsKey(korean))
                KoreanToEnglish.Add(korean, english);
        }

        private static string SettingsDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointLite");
            }
        }

        private static string SettingsPath
        {
            get { return Path.Combine(SettingsDirectory, "settings.ini"); }
        }

        private static AppLanguage LoadLanguage()
        {
            try
            {
                if (!File.Exists(SettingsPath))
                    return AppLanguage.Korean;

                string[] lines = File.ReadAllLines(SettingsPath, Encoding.UTF8);

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();

                    if (string.Equals(line, "language=en-US", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(line, "language=en", StringComparison.OrdinalIgnoreCase))
                    {
                        return AppLanguage.English;
                    }

                    if (string.Equals(line, "language=ko-KR", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(line, "language=ko", StringComparison.OrdinalIgnoreCase))
                    {
                        return AppLanguage.Korean;
                    }
                }
            }
            catch
            {
            }

            return AppLanguage.Korean;
        }

        private static void SaveLanguage(AppLanguage language)
        {
            try
            {
                Directory.CreateDirectory(SettingsDirectory);
                File.WriteAllText(
                    SettingsPath,
                    language == AppLanguage.English
                        ? "language=en-US\r\n"
                        : "language=ko-KR\r\n",
                    Encoding.UTF8);
            }
            catch
            {
            }
        }
    }
}
