using System;
using System.Collections.Generic;
using System.Text;

namespace PptxViewer
{
    internal enum RichTextFormatField
    {
        FontFamily,
        FontSize,
        Bold,
        Italic,
        Underline,
        Color
    }

    internal sealed class RichTextFormatChange
    {
        public RichTextFormatField Field { get; set; }
        public string FontFamily { get; set; }
        public float FontSizePoints { get; set; }
        public bool BooleanValue { get; set; }
        public string ColorHex { get; set; }
    }

    internal sealed class RichTextSelectionStyle
    {
        public string FontFamily { get; set; }
        public float FontSizePoints { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string ColorHex { get; set; }

        public bool FontFamilyMixed { get; set; }
        public bool FontSizeMixed { get; set; }
        public bool BoldMixed { get; set; }
        public bool ItalicMixed { get; set; }
        public bool UnderlineMixed { get; set; }
        public bool ColorMixed { get; set; }
    }

    internal sealed class RichTextEditorSegment
    {
        public int Start { get; set; }
        public int Length { get; set; }
        public PresentationTextRun Run { get; set; }
    }

    internal static class RichTextSelectionEditor
    {
        public static string GetEditorText(
            PresentationTextBox box)
        {
            if (box == null)
                return string.Empty;

            if (!box.HasRichText)
                return NormalizeForEditor(box.Text);

            StringBuilder builder =
                new StringBuilder();

            for (int p = 0;
                 p < box.RichParagraphs.Count;
                 p++)
            {
                if (p > 0)
                    builder.Append('\n');

                PresentationTextParagraph paragraph =
                    box.RichParagraphs[p];

                if (paragraph != null)
                    builder.Append(
                        NormalizeForEditor(
                            paragraph.GetPlainText()));
            }

            return builder.ToString();
        }

        public static string ToModelText(
            string editorText)
        {
            string normalized =
                NormalizeForEditor(
                    editorText);

            if (Environment.NewLine == "\n")
                return normalized;

            return normalized.Replace(
                "\n",
                Environment.NewLine);
        }

        public static List<RichTextEditorSegment>
            GetSegments(
                PresentationTextBox box)
        {
            List<RichTextEditorSegment> result =
                new List<RichTextEditorSegment>();

            if (box == null ||
                !box.HasRichText)
            {
                return result;
            }

            int cursor = 0;

            for (int p = 0;
                 p < box.RichParagraphs.Count;
                 p++)
            {
                PresentationTextParagraph paragraph =
                    box.RichParagraphs[p];

                if (paragraph != null)
                {
                    for (int r = 0;
                         r < paragraph.Runs.Count;
                         r++)
                    {
                        PresentationTextRun run =
                            paragraph.Runs[r];

                        if (run == null)
                            continue;

                        string text =
                            NormalizeForEditor(
                                run.Text);

                        if (text.Length > 0)
                        {
                            RichTextEditorSegment segment =
                                new RichTextEditorSegment();
                            segment.Start = cursor;
                            segment.Length = text.Length;
                            segment.Run = run;
                            result.Add(segment);
                        }

                        cursor += text.Length;
                    }
                }

                if (p <
                    box.RichParagraphs.Count - 1)
                {
                    cursor++;
                }
            }

            return result;
        }

        public static bool ApplySelection(
            PresentationTextBox box,
            int selectionStart,
            int selectionLength,
            RichTextFormatChange change)
        {
            if (box == null ||
                change == null ||
                selectionLength <= 0)
            {
                return false;
            }

            string editorText =
                GetEditorText(box);

            int start =
                Math.Max(
                    0,
                    Math.Min(
                        editorText.Length,
                        selectionStart));
            int end =
                Math.Max(
                    start,
                    Math.Min(
                        editorText.Length,
                        start +
                        selectionLength));

            if (end <= start)
                return false;

            EnsureRichText(box);

            bool changed = false;
            int paragraphStart = 0;

            for (int p = 0;
                 p < box.RichParagraphs.Count;
                 p++)
            {
                PresentationTextParagraph paragraph =
                    box.RichParagraphs[p];

                if (paragraph == null)
                {
                    paragraphStart++;
                    continue;
                }

                List<PresentationTextRun> rebuilt =
                    new List<PresentationTextRun>();
                int runStart =
                    paragraphStart;

                for (int r = 0;
                     r < paragraph.Runs.Count;
                     r++)
                {
                    PresentationTextRun run =
                        paragraph.Runs[r];

                    if (run == null)
                        continue;

                    string runText =
                        NormalizeForEditor(
                            run.Text);
                    int runLength =
                        runText.Length;
                    int runEnd =
                        runStart +
                        runLength;

                    int overlapStart =
                        Math.Max(
                            start,
                            runStart);
                    int overlapEnd =
                        Math.Min(
                            end,
                            runEnd);

                    if (runLength == 0 ||
                        overlapEnd <=
                        overlapStart)
                    {
                        rebuilt.Add(
                            run.Clone());
                    }
                    else
                    {
                        int localStart =
                            overlapStart -
                            runStart;
                        int localEnd =
                            overlapEnd -
                            runStart;

                        if (localStart > 0)
                        {
                            PresentationTextRun prefix =
                                run.Clone();
                            prefix.Text =
                                runText.Substring(
                                    0,
                                    localStart);
                            rebuilt.Add(prefix);
                        }

                        PresentationTextRun selected =
                            run.Clone();
                        selected.Text =
                            runText.Substring(
                                localStart,
                                localEnd -
                                localStart);
                        ApplyChange(
                            selected,
                            change);
                        rebuilt.Add(selected);
                        changed = true;

                        if (localEnd <
                            runLength)
                        {
                            PresentationTextRun suffix =
                                run.Clone();
                            suffix.Text =
                                runText.Substring(
                                    localEnd);
                            rebuilt.Add(suffix);
                        }
                    }

                    runStart = runEnd;
                }

                paragraph.Runs.Clear();

                for (int r = 0;
                     r < rebuilt.Count;
                     r++)
                {
                    paragraph.Runs.Add(
                        rebuilt[r]);
                }

                MergeAdjacentRuns(paragraph);
                paragraphStart =
                    runStart + 1;
            }

            if (changed)
                box.RefreshTextFromRichParagraphs();

            return changed;
        }

        public static bool TryGetSelectionStyle(
            PresentationTextBox box,
            int selectionStart,
            int selectionLength,
            out RichTextSelectionStyle style)
        {
            style = null;

            if (box == null)
                return false;

            string text =
                GetEditorText(box);

            if (text.Length == 0)
            {
                style =
                    FromBoxDefaults(box);
                return true;
            }

            int start =
                Math.Max(
                    0,
                    Math.Min(
                        text.Length,
                        selectionStart));
            int length =
                Math.Max(
                    0,
                    selectionLength);
            int end;

            if (length == 0)
            {
                if (start >= text.Length)
                {
                    start =
                        Math.Max(
                            0,
                            text.Length - 1);
                }

                end =
                    Math.Min(
                        text.Length,
                        start + 1);
            }
            else
            {
                end =
                    Math.Min(
                        text.Length,
                        start +
                        length);
            }

            if (!box.HasRichText)
            {
                style =
                    FromBoxDefaults(box);
                return true;
            }

            List<RichTextEditorSegment> segments =
                GetSegments(box);
            PresentationTextRun first = null;

            for (int i = 0;
                 i < segments.Count;
                 i++)
            {
                RichTextEditorSegment segment =
                    segments[i];
                int segmentEnd =
                    segment.Start +
                    segment.Length;

                if (segmentEnd <= start ||
                    segment.Start >= end)
                {
                    continue;
                }

                if (first == null)
                {
                    first = segment.Run;
                    style =
                        FromRun(
                            segment.Run,
                            box);
                }
                else
                {
                    MarkDifferences(
                        style,
                        first,
                        segment.Run,
                        box);
                }
            }

            if (style == null)
                style = FromBoxDefaults(box);

            return true;
        }

        public static void ApplyWholeBoxDefaults(
            PresentationTextBox box,
            RichTextFormatChange change)
        {
            if (box == null ||
                change == null)
            {
                return;
            }

            if (change.Field ==
                RichTextFormatField.FontFamily)
            {
                if (!string.IsNullOrEmpty(
                        change.FontFamily))
                {
                    box.FontFamily =
                        change.FontFamily;
                }
            }
            else if (change.Field ==
                     RichTextFormatField.FontSize)
            {
                box.FontSizePoints =
                    Math.Max(
                        1f,
                        Math.Min(
                            400f,
                            change.FontSizePoints));
            }
            else if (change.Field ==
                     RichTextFormatField.Bold)
            {
                box.Bold =
                    change.BooleanValue;
            }
            else if (change.Field ==
                     RichTextFormatField.Italic)
            {
                box.Italic =
                    change.BooleanValue;
            }
            else if (change.Field ==
                     RichTextFormatField.Color)
            {
                if (!string.IsNullOrEmpty(
                        change.ColorHex))
                {
                    box.ColorHex =
                        change.ColorHex;
                }
            }
        }

        public static void ApplyParagraphAlignment(
            PresentationTextBox box,
            PresentationTextAlignment alignment)
        {
            if (box == null)
                return;

            box.Alignment = alignment;

            for (int p = 0;
                 p < box.RichParagraphs.Count;
                 p++)
            {
                PresentationTextParagraph paragraph =
                    box.RichParagraphs[p];

                if (paragraph != null)
                    paragraph.Alignment =
                        alignment;
            }
        }

        private static void EnsureRichText(
            PresentationTextBox box)
        {
            if (box.HasRichText)
                return;

            string editorText =
                NormalizeForEditor(
                    box.Text);
            string[] paragraphs =
                editorText.Split(
                    new char[] { '\n' });
            List<PresentationTextParagraph> result =
                new List<PresentationTextParagraph>();

            for (int i = 0;
                 i < paragraphs.Length;
                 i++)
            {
                PresentationTextParagraph paragraph =
                    new PresentationTextParagraph();
                paragraph.Alignment =
                    box.Alignment;

                PresentationTextRun run =
                    new PresentationTextRun();
                run.Text =
                    paragraphs[i] ??
                    string.Empty;
                run.FontFamily =
                    box.FontFamily;
                run.FontSizePoints =
                    box.FontSizePoints;
                run.Bold =
                    box.Bold;
                run.Italic =
                    box.Italic;
                run.Underline =
                    false;
                run.ColorHex =
                    box.ColorHex;
                paragraph.Runs.Add(run);
                result.Add(paragraph);
            }

            if (result.Count == 0)
                result.Add(
                    new PresentationTextParagraph());

            box.SetRichParagraphs(result);
        }

        private static void ApplyChange(
            PresentationTextRun run,
            RichTextFormatChange change)
        {
            if (run == null ||
                change == null)
            {
                return;
            }

            if (change.Field ==
                RichTextFormatField.FontFamily)
            {
                if (!string.IsNullOrEmpty(
                        change.FontFamily))
                {
                    run.FontFamily =
                        change.FontFamily;
                }
            }
            else if (change.Field ==
                     RichTextFormatField.FontSize)
            {
                run.FontSizePoints =
                    Math.Max(
                        1f,
                        Math.Min(
                            400f,
                            change.FontSizePoints));
            }
            else if (change.Field ==
                     RichTextFormatField.Bold)
            {
                run.Bold =
                    change.BooleanValue;
            }
            else if (change.Field ==
                     RichTextFormatField.Italic)
            {
                run.Italic =
                    change.BooleanValue;
            }
            else if (change.Field ==
                     RichTextFormatField.Underline)
            {
                run.Underline =
                    change.BooleanValue;
            }
            else if (change.Field ==
                     RichTextFormatField.Color)
            {
                if (!string.IsNullOrEmpty(
                        change.ColorHex))
                {
                    run.ColorHex =
                        change.ColorHex;
                }
            }
        }

        private static void MergeAdjacentRuns(
            PresentationTextParagraph paragraph)
        {
            if (paragraph == null ||
                paragraph.Runs.Count < 2)
            {
                return;
            }

            List<PresentationTextRun> merged =
                new List<PresentationTextRun>();

            for (int i = 0;
                 i < paragraph.Runs.Count;
                 i++)
            {
                PresentationTextRun current =
                    paragraph.Runs[i];

                if (current == null)
                    continue;

                if (merged.Count > 0 &&
                    SameFormatting(
                        merged[
                            merged.Count - 1],
                        current))
                {
                    merged[
                        merged.Count - 1].Text =
                        (merged[
                            merged.Count - 1].Text ??
                         string.Empty) +
                        (current.Text ??
                         string.Empty);
                }
                else
                {
                    merged.Add(
                        current.Clone());
                }
            }

            paragraph.Runs.Clear();

            for (int i = 0;
                 i < merged.Count;
                 i++)
            {
                paragraph.Runs.Add(
                    merged[i]);
            }
        }

        private static bool SameFormatting(
            PresentationTextRun left,
            PresentationTextRun right)
        {
            if (left == null ||
                right == null)
            {
                return false;
            }

            return
                string.Equals(
                    left.FontFamily,
                    right.FontFamily,
                    StringComparison.OrdinalIgnoreCase) &&
                Math.Abs(
                    left.FontSizePoints -
                    right.FontSizePoints) <
                    0.01f &&
                left.Bold == right.Bold &&
                left.Italic == right.Italic &&
                left.Underline == right.Underline &&
                string.Equals(
                    NormalizeColor(
                        left.ColorHex),
                    NormalizeColor(
                        right.ColorHex),
                    StringComparison.OrdinalIgnoreCase) &&
                left.BaselinePercent ==
                    right.BaselinePercent;
        }

        private static RichTextSelectionStyle FromBoxDefaults(
            PresentationTextBox box)
        {
            RichTextSelectionStyle style =
                new RichTextSelectionStyle();
            style.FontFamily =
                box.FontFamily;
            style.FontSizePoints =
                box.FontSizePoints;
            style.Bold =
                box.Bold;
            style.Italic =
                box.Italic;
            style.Underline =
                false;
            style.ColorHex =
                box.ColorHex;
            return style;
        }

        private static RichTextSelectionStyle FromRun(
            PresentationTextRun run,
            PresentationTextBox box)
        {
            RichTextSelectionStyle style =
                new RichTextSelectionStyle();
            style.FontFamily =
                string.IsNullOrEmpty(
                    run.FontFamily)
                    ? box.FontFamily
                    : run.FontFamily;
            style.FontSizePoints =
                run.FontSizePoints > 0f
                    ? run.FontSizePoints
                    : box.FontSizePoints;
            style.Bold =
                run.Bold;
            style.Italic =
                run.Italic;
            style.Underline =
                run.Underline;
            style.ColorHex =
                string.IsNullOrEmpty(
                    run.ColorHex)
                    ? box.ColorHex
                    : run.ColorHex;
            return style;
        }

        private static void MarkDifferences(
            RichTextSelectionStyle style,
            PresentationTextRun first,
            PresentationTextRun current,
            PresentationTextBox box)
        {
            string firstFamily =
                string.IsNullOrEmpty(
                    first.FontFamily)
                    ? box.FontFamily
                    : first.FontFamily;
            string currentFamily =
                string.IsNullOrEmpty(
                    current.FontFamily)
                    ? box.FontFamily
                    : current.FontFamily;
            float firstSize =
                first.FontSizePoints > 0f
                    ? first.FontSizePoints
                    : box.FontSizePoints;
            float currentSize =
                current.FontSizePoints > 0f
                    ? current.FontSizePoints
                    : box.FontSizePoints;
            string firstColor =
                string.IsNullOrEmpty(
                    first.ColorHex)
                    ? box.ColorHex
                    : first.ColorHex;
            string currentColor =
                string.IsNullOrEmpty(
                    current.ColorHex)
                    ? box.ColorHex
                    : current.ColorHex;

            if (!string.Equals(
                    firstFamily,
                    currentFamily,
                    StringComparison.OrdinalIgnoreCase))
            {
                style.FontFamilyMixed = true;
            }

            if (Math.Abs(
                    firstSize -
                    currentSize) >= 0.01f)
            {
                style.FontSizeMixed = true;
            }

            if (first.Bold != current.Bold)
                style.BoldMixed = true;

            if (first.Italic != current.Italic)
                style.ItalicMixed = true;

            if (first.Underline !=
                current.Underline)
            {
                style.UnderlineMixed = true;
            }

            if (!string.Equals(
                    NormalizeColor(firstColor),
                    NormalizeColor(currentColor),
                    StringComparison.OrdinalIgnoreCase))
            {
                style.ColorMixed = true;
            }
        }

        private static string NormalizeForEditor(
            string value)
        {
            return (value ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");
        }

        private static string NormalizeColor(
            string value)
        {
            string candidate =
                (value ?? string.Empty)
                .Trim()
                .TrimStart('#')
                .ToUpperInvariant();

            return candidate.Length == 6
                ? candidate
                : "20242A";
        }
    }

    internal static class RichTextSelectionEditorDiagnostics
    {
        public static void Validate()
        {
            PresentationTextBox box =
                new PresentationTextBox();
            box.Text =
                "Alpha Beta Gamma";
            box.FontFamily =
                "Arial";
            box.FontSizePoints =
                20f;
            box.ColorHex =
                "20242A";

            int betaStart =
                RichTextSelectionEditor
                    .GetEditorText(box)
                    .IndexOf(
                        "Beta",
                        StringComparison.Ordinal);

            RichTextFormatChange bold =
                new RichTextFormatChange();
            bold.Field =
                RichTextFormatField.Bold;
            bold.BooleanValue = true;

            if (!RichTextSelectionEditor.ApplySelection(
                    box,
                    betaStart,
                    4,
                    bold))
            {
                throw new InvalidOperationException(
                    "Rich-text selection formatting did not change the target range.");
            }

            RichTextFormatChange underline =
                new RichTextFormatChange();
            underline.Field =
                RichTextFormatField.Underline;
            underline.BooleanValue = true;

            RichTextSelectionEditor.ApplySelection(
                box,
                betaStart,
                4,
                underline);

            if (!box.HasRichText ||
                box.RichParagraphs.Count != 1 ||
                box.RichParagraphs[0].Runs.Count != 3 ||
                box.RichParagraphs[0].Runs[1].Text !=
                    "Beta" ||
                !box.RichParagraphs[0].Runs[1].Bold ||
                !box.RichParagraphs[0].Runs[1].Underline ||
                box.RichParagraphs[0].Runs[0].Bold ||
                box.RichParagraphs[0].Runs[2].Underline)
            {
                throw new InvalidOperationException(
                    "Rich-text selection formatting did not preserve surrounding runs.");
            }

            if (box.Text !=
                "Alpha Beta Gamma")
            {
                throw new InvalidOperationException(
                    "Rich-text selection formatting changed the plain text.");
            }

            RichTextSelectionStyle style;

            if (!RichTextSelectionEditor.TryGetSelectionStyle(
                    box,
                    betaStart,
                    4,
                    out style) ||
                style == null ||
                !style.Bold ||
                !style.Underline)
            {
                throw new InvalidOperationException(
                    "Rich-text selection style inspection failed.");
            }

            RichTextFormatChange unbold =
                new RichTextFormatChange();
            unbold.Field =
                RichTextFormatField.Bold;
            unbold.BooleanValue = false;
            RichTextSelectionEditor.ApplySelection(
                box,
                betaStart,
                4,
                unbold);

            RichTextFormatChange noUnderline =
                new RichTextFormatChange();
            noUnderline.Field =
                RichTextFormatField.Underline;
            noUnderline.BooleanValue = false;
            RichTextSelectionEditor.ApplySelection(
                box,
                betaStart,
                4,
                noUnderline);

            if (box.RichParagraphs[0].Runs.Count != 1 ||
                box.RichParagraphs[0].Runs[0].Text !=
                    "Alpha Beta Gamma")
            {
                throw new InvalidOperationException(
                    "Equivalent adjacent rich-text runs were not merged.");
            }

            PresentationTextBox multiLine =
                new PresentationTextBox();
            multiLine.Text =
                RichTextSelectionEditor.ToModelText(
                    "One\nTwo\nThree");

            int twoStart =
                RichTextSelectionEditor
                    .GetEditorText(multiLine)
                    .IndexOf(
                        "Two",
                        StringComparison.Ordinal);

            RichTextFormatChange color =
                new RichTextFormatChange();
            color.Field =
                RichTextFormatField.Color;
            color.ColorHex =
                "AA3300";

            RichTextSelectionEditor.ApplySelection(
                multiLine,
                twoStart,
                3,
                color);

            if (multiLine.RichParagraphs.Count != 3 ||
                multiLine.RichParagraphs[1].Runs.Count != 1 ||
                multiLine.RichParagraphs[1].Runs[0].ColorHex !=
                    "AA3300" ||
                RichTextSelectionEditor.GetEditorText(
                    multiLine) !=
                    "One\nTwo\nThree")
            {
                throw new InvalidOperationException(
                    "Rich-text selection formatting failed across paragraph mapping.");
            }
        }
    }
}
