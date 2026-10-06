using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal static partial class InternalPptxRenderer
    {
        private sealed class RichRunStyle
        {
            public float SizePt = 18f;
            public string FontName = "Arial";
            public Color Color = Color.Black;
            public FontStyle FontStyle = FontStyle.Regular;

            public RichRunStyle Clone()
            {
                return new RichRunStyle
                {
                    SizePt = SizePt,
                    FontName = FontName,
                    Color = Color,
                    FontStyle = FontStyle
                };
            }
        }

        private sealed class RichTextToken
        {
            public string Text;
            public RichRunStyle Style;
            public bool Break;
            public bool Tab;
        }

        private sealed class RichTabStop
        {
            public float Position;
            public string Alignment = "l";
        }

        private sealed class RichTextLine
        {
            public readonly List<RichTextToken> Tokens = new List<RichTextToken>();
            public readonly List<RichTabStop> TabStops = new List<RichTabStop>();
            public float Width;
            public float Height;
            public float LeftOffset;
            public float AvailableWidth;
            public float DefaultTabSize = 48f;
            public bool RightToLeft;
        }

        private sealed class RichParagraphLayout
        {
            public readonly List<RichTextLine> Lines = new List<RichTextLine>();
            public readonly List<RichTabStop> TabStops = new List<RichTabStop>();
            public float Before;
            public float After;
            public float TotalHeight;
            public float LineFactor = 1f;
            public float AbsoluteLineHeight;
            public float DefaultTabSize = 48f;
            public bool RightToLeft;
            public StringAlignment Alignment = StringAlignment.Near;
        }

        private static void DrawRichShapeText(
            Graphics g,
            XmlNode shape,
            RectangleF rect,
            Dictionary<string, Color> theme,
            int slideNumber)
        {
            XmlNode txBody = DirectChild(shape, "txBody");
            if (txBody == null)
                return;

            XmlNode bodyPr = DirectChild(txBody, "bodyPr");
            float lIns = 6f;
            float rIns = 6f;
            float tIns = 4f;
            float bIns = 4f;
            bool wrap = true;
            bool normalAutoFitEnabled = false;
            bool shapeAutoFitEnabled = false;
            float fontScale = 1f;
            float lineReduction = 0f;
            float defaultTabSize = 48f;
            string anchor = "t";
            string vertical = "";

            if (bodyPr != null)
            {
                lIns = EmuToRenderPixels(GetLong(bodyPr, "lIns", 45720));
                rIns = EmuToRenderPixels(GetLong(bodyPr, "rIns", 45720));
                tIns = EmuToRenderPixels(GetLong(bodyPr, "tIns", 22860));
                bIns = EmuToRenderPixels(GetLong(bodyPr, "bIns", 22860));

                string wrapValue = GetAttr(bodyPr, "wrap");
                wrap = !string.Equals(wrapValue, "none", StringComparison.OrdinalIgnoreCase);

                string bodyAnchor = GetAttr(bodyPr, "anchor");
                if (!string.IsNullOrEmpty(bodyAnchor))
                    anchor = bodyAnchor;

                vertical = GetAttr(bodyPr, "vert");

                long tabSizeEmu =
                    GetLong(
                        bodyPr,
                        "defTabSz",
                        914400);
                defaultTabSize =
                    Math.Max(
                        8f,
                        EmuToRenderPixels(
                            tabSizeEmu));

                XmlNode normalAutoFit = DirectChild(bodyPr, "normAutofit");
                XmlNode shapeAutoFit = DirectChild(bodyPr, "spAutoFit");
                XmlNode noAutoFit = DirectChild(bodyPr, "noAutofit");

                if (normalAutoFit != null)
                {
                    normalAutoFitEnabled = true;
                    long scale = GetLong(normalAutoFit, "fontScale", 100000);
                    fontScale = Math.Max(0.1f, Math.Min(1f, scale / 100000f));

                    long reduction = GetLong(normalAutoFit, "lnSpcReduction", 0);
                    lineReduction = Math.Max(0f, Math.Min(0.8f, reduction / 100000f));
                }
                else if (shapeAutoFit != null)
                {
                    shapeAutoFitEnabled = true;
                }
                else if (noAutoFit != null)
                {
                    fontScale = 1f;
                    lineReduction = 0f;
                }
            }

            RectangleF inner = new RectangleF(
                rect.X + lIns,
                rect.Y + tIns,
                Math.Max(1f, rect.Width - lIns - rIns),
                Math.Max(1f, rect.Height - tIns - bIns));

            GraphicsState textState = g.Save();

            try
            {
                RectangleF layoutRect = inner;

                if (string.Equals(vertical, "vert", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(vertical, "eaVert", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(vertical, "vert270", StringComparison.OrdinalIgnoreCase))
                {
                    float cx = inner.Left + inner.Width / 2f;
                    float cy = inner.Top + inner.Height / 2f;
                    float angle = string.Equals(vertical, "vert270", StringComparison.OrdinalIgnoreCase)
                        ? -90f
                        : 90f;

                    g.TranslateTransform(cx, cy);
                    g.RotateTransform(angle);
                    g.TranslateTransform(-cx, -cy);

                    layoutRect = new RectangleF(
                        cx - inner.Height / 2f,
                        cy - inner.Width / 2f,
                        inner.Height,
                        inner.Width);
                }

                float totalHeight;
                List<RichParagraphLayout> layouts =
                    BuildRichParagraphLayouts(
                        g,
                        txBody,
                        layoutRect,
                        theme,
                        slideNumber,
                        wrap,
                        fontScale,
                        lineReduction,
                        defaultTabSize,
                        out totalHeight);

                if (normalAutoFitEnabled &&
                    totalHeight > layoutRect.Height + 1f &&
                    fontScale > 0.11f)
                {
                    float fitRatio =
                        Math.Max(
                            0.1f,
                            Math.Min(
                                1f,
                                layoutRect.Height /
                                Math.Max(
                                    1f,
                                    totalHeight)));

                    float adjustedScale =
                        Math.Max(
                            0.1f,
                            Math.Min(
                                fontScale,
                                fontScale *
                                fitRatio *
                                0.98f));

                    if (adjustedScale <
                        fontScale - 0.005f)
                    {
                        layouts =
                            BuildRichParagraphLayouts(
                                g,
                                txBody,
                                layoutRect,
                                theme,
                                slideNumber,
                                wrap,
                                adjustedScale,
                                lineReduction,
                                defaultTabSize,
                                out totalHeight);
                    }
                }

                if (shapeAutoFitEnabled &&
                    totalHeight > layoutRect.Height)
                {
                    layoutRect.Height =
                        totalHeight;
                }

                float y = layoutRect.Top;
                if (anchor == "ctr")
                    y += Math.Max(0f, (layoutRect.Height - totalHeight) / 2f);
                else if (anchor == "b")
                    y += Math.Max(0f, layoutRect.Height - totalHeight);

                g.SetClip(layoutRect, CombineMode.Intersect);

                for (int i = 0; i < layouts.Count; i++)
                {
                    RichParagraphLayout paragraph = layouts[i];
                    y += paragraph.Before;

                    for (int lineIndex = 0; lineIndex < paragraph.Lines.Count; lineIndex++)
                    {
                        RichTextLine line = paragraph.Lines[lineIndex];
                        float x = layoutRect.Left + line.LeftOffset;

                        if (paragraph.Alignment == StringAlignment.Center)
                            x += Math.Max(0f, (line.AvailableWidth - line.Width) / 2f);
                        else if (paragraph.Alignment == StringAlignment.Far)
                            x += Math.Max(0f, line.AvailableWidth - line.Width);

                        DrawRichTextLine(g, line, x, y);

                        float advance = line.Height;
                        if (paragraph.AbsoluteLineHeight > 0f)
                            advance = Math.Max(advance, paragraph.AbsoluteLineHeight);
                        else
                            advance *= paragraph.LineFactor;

                        y += Math.Max(1f, advance);
                    }

                    y += paragraph.After;

                    if (y > layoutRect.Bottom + 1f)
                        break;
                }
            }
            finally
            {
                g.Restore(textState);
            }
        }

        private static List<RichParagraphLayout> BuildRichParagraphLayouts(
            Graphics g,
            XmlNode txBody,
            RectangleF layoutRect,
            Dictionary<string, Color> theme,
            int slideNumber,
            bool wrap,
            float fontScale,
            float lineReduction,
            float defaultTabSize,
            out float totalHeight)
        {
            List<RichParagraphLayout> layouts =
                new List<RichParagraphLayout>();

            totalHeight = 0f;
            int paragraphIndex = 0;

            foreach (XmlNode paragraph in txBody.ChildNodes)
            {
                if (paragraph.LocalName != "p")
                    continue;

                RichParagraphLayout paragraphLayout =
                    BuildRichParagraphLayout(
                        g,
                        paragraph,
                        layoutRect,
                        theme,
                        slideNumber,
                        paragraphIndex,
                        wrap,
                        fontScale,
                        lineReduction,
                        defaultTabSize);

                layouts.Add(paragraphLayout);
                totalHeight += paragraphLayout.TotalHeight;
                paragraphIndex++;
            }

            return layouts;
        }

        private static RichParagraphLayout BuildRichParagraphLayout(
            Graphics g,
            XmlNode paragraph,
            RectangleF inner,
            Dictionary<string, Color> theme,
            int slideNumber,
            int paragraphIndex,
            bool wrap,
            float fontScale,
            float lineReduction,
            float defaultTabSize)
        {
            RichParagraphLayout result = new RichParagraphLayout();
            result.DefaultTabSize = Math.Max(8f, defaultTabSize);
            XmlNode pPr = DirectChild(paragraph, "pPr");
            XmlNode defaultRPr = pPr != null ? DirectChild(pPr, "defRPr") : null;
            XmlNode endRPr = DirectChild(paragraph, "endParaRPr");

            RichRunStyle rawBaseStyle = new RichRunStyle();
            rawBaseStyle = ApplyRichRunProperties(
                rawBaseStyle,
                defaultRPr != null ? defaultRPr : endRPr,
                theme,
                "",
                1f);

            RichRunStyle displayBaseStyle = ApplyRichRunProperties(
                rawBaseStyle,
                null,
                theme,
                "",
                fontScale);

            float marginLeft = 0f;
            float marginRight = 0f;
            float indent = 0f;

            if (pPr != null)
            {
                marginLeft = EmuToRenderPixels(GetLong(pPr, "marL", 0));
                marginRight = EmuToRenderPixels(GetLong(pPr, "marR", 0));
                indent = EmuToRenderPixels(GetLong(pPr, "indent", 0));

                string rtlValue = GetAttr(pPr, "rtl");
                result.RightToLeft =
                    IsRichBooleanTrue(rtlValue);

                string alignment = GetAttr(pPr, "algn");
                if (alignment == "ctr")
                    result.Alignment = StringAlignment.Center;
                else if (alignment == "r")
                    result.Alignment = StringAlignment.Far;
                else if (alignment == "l")
                    result.Alignment = StringAlignment.Near;
                else if (result.RightToLeft)
                    result.Alignment = StringAlignment.Far;

                XmlNode tabList =
                    DirectChild(
                        pPr,
                        "tabLst");

                if (tabList != null)
                {
                    foreach (XmlNode tab in tabList.ChildNodes)
                    {
                        if (tab.LocalName != "tab")
                            continue;

                        float position =
                            EmuToRenderPixels(
                                GetLong(
                                    tab,
                                    "pos",
                                    0));

                        if (position > 0f)
                        {
                            RichTabStop stop =
                                new RichTabStop();

                            stop.Position =
                                position;

                            string tabAlignment =
                                GetAttr(
                                    tab,
                                    "algn");

                            if (!string.IsNullOrEmpty(
                                    tabAlignment))
                            {
                                stop.Alignment =
                                    tabAlignment;
                            }

                            result.TabStops.Add(
                                stop);
                        }
                    }

                    result.TabStops.Sort(
                        delegate(
                            RichTabStop left,
                            RichTabStop right)
                        {
                            return left.Position.CompareTo(
                                right.Position);
                        });
                }

                result.Before = ReadRichSpacingPixels(g, DirectChild(pPr, "spcBef"), displayBaseStyle.SizePt);
                result.After = ReadRichSpacingPixels(g, DirectChild(pPr, "spcAft"), displayBaseStyle.SizePt);
                ReadRichLineSpacing(
                    g,
                    DirectChild(pPr, "lnSpc"),
                    displayBaseStyle.SizePt,
                    lineReduction,
                    result);
            }

            List<RichTextToken> tokens = new List<RichTextToken>();
            AddRichBulletToken(tokens, pPr, displayBaseStyle, paragraphIndex);

            foreach (XmlNode child in paragraph.ChildNodes)
            {
                if (child.LocalName == "br")
                {
                    tokens.Add(new RichTextToken { Break = true });
                    continue;
                }

                if (child.LocalName == "tab")
                {
                    tokens.Add(new RichTextToken
                    {
                        Tab = true,
                        Style = displayBaseStyle.Clone()
                    });
                    continue;
                }

                if (child.LocalName != "r" && child.LocalName != "fld")
                    continue;

                string text = "";
                foreach (XmlNode t in FindAll(child, "t"))
                {
                    if (child.LocalName == "fld" &&
                        string.Equals(GetAttr(child, "type"), "slidenum", StringComparison.OrdinalIgnoreCase))
                    {
                        text += slideNumber.ToString();
                    }
                    else
                    {
                        text += t.InnerText;
                    }
                }

                if (text.Length == 0)
                    continue;

                XmlNode rPr = DirectChild(child, "rPr");
                RichRunStyle style = ApplyRichRunProperties(
                    rawBaseStyle,
                    rPr,
                    theme,
                    text,
                    fontScale);

                AddRichTextTokens(tokens, text, style);
            }

            if (tokens.Count == 0)
            {
                RichTextLine blank = new RichTextLine();
                blank.Height = RichFontHeight(g, displayBaseStyle);
                blank.LeftOffset = Math.Max(0f, marginLeft + indent);
                blank.AvailableWidth = Math.Max(1f, inner.Width - blank.LeftOffset - marginRight);
                result.Lines.Add(blank);
                result.TotalHeight = result.Before + blank.Height + result.After;
                return result;
            }

            RichTextLine line = CreateRichLine(
                inner,
                marginLeft,
                marginRight,
                indent,
                true,
                displayBaseStyle,
                g,
                result.TabStops,
                result.DefaultTabSize,
                result.RightToLeft);

            for (int i = 0; i < tokens.Count; i++)
            {
                RichTextToken token = tokens[i];

                if (token.Break)
                {
                    TrimTrailingRichWhitespace(g, line);
                    result.Lines.Add(line);
                    line = CreateRichLine(
                        inner,
                        marginLeft,
                        marginRight,
                        indent,
                        false,
                        displayBaseStyle,
                        g,
                        result.TabStops,
                        result.DefaultTabSize,
                        result.RightToLeft);
                    continue;
                }

                float tokenWidth;
                float tokenHeight;

                if (token.Tab)
                {
                    tokenWidth =
                        CalculateRichTabAdvance(
                            line,
                            line.LeftOffset +
                            line.Width);
                    tokenHeight =
                        RichFontHeight(
                            g,
                            token.Style != null
                                ? token.Style
                                : displayBaseStyle);
                }
                else
                {
                    MeasureRichToken(
                        g,
                        token,
                        out tokenWidth,
                        out tokenHeight);
                }

                bool whitespaceOnly =
                    !token.Tab &&
                    string.IsNullOrWhiteSpace(token.Text);
                bool wouldOverflow =
                    wrap &&
                    line.Tokens.Count > 0 &&
                    !whitespaceOnly &&
                    line.Width + tokenWidth > line.AvailableWidth;

                if (wouldOverflow)
                {
                    TrimTrailingRichWhitespace(g, line);
                    result.Lines.Add(line);
                    line = CreateRichLine(
                        inner,
                        marginLeft,
                        marginRight,
                        indent,
                        false,
                        displayBaseStyle,
                        g,
                        result.TabStops,
                        result.DefaultTabSize,
                        result.RightToLeft);

                    if (whitespaceOnly)
                        continue;
                }

                if (line.Tokens.Count == 0 && whitespaceOnly)
                    continue;

                line.Tokens.Add(token);
                line.Width += tokenWidth;
                line.Height = Math.Max(line.Height, tokenHeight);
            }

            TrimTrailingRichWhitespace(g, line);
            if (line.Tokens.Count > 0 || result.Lines.Count == 0)
                result.Lines.Add(line);

            float lineHeightTotal = 0f;
            for (int i = 0; i < result.Lines.Count; i++)
            {
                RichTextLine currentLine = result.Lines[i];
                float advance = currentLine.Height;

                if (result.AbsoluteLineHeight > 0f)
                    advance = Math.Max(advance, result.AbsoluteLineHeight);
                else
                    advance *= result.LineFactor;

                lineHeightTotal += Math.Max(1f, advance);
            }

            result.TotalHeight = result.Before + lineHeightTotal + result.After;
            return result;
        }

        private static RichTextLine CreateRichLine(
            RectangleF inner,
            float marginLeft,
            float marginRight,
            float indent,
            bool firstLine,
            RichRunStyle baseStyle,
            Graphics g,
            List<RichTabStop> tabStops,
            float defaultTabSize,
            bool rightToLeft)
        {
            RichTextLine line =
                new RichTextLine();

            float offset =
                marginLeft +
                (firstLine
                    ? indent
                    : 0f);

            offset =
                Math.Max(
                    0f,
                    offset);

            line.LeftOffset = offset;
            line.AvailableWidth =
                Math.Max(
                    1f,
                    inner.Width -
                    offset -
                    Math.Max(
                        0f,
                        marginRight));
            line.Height =
                RichFontHeight(
                    g,
                    baseStyle);
            line.DefaultTabSize =
                Math.Max(
                    8f,
                    defaultTabSize);
            line.RightToLeft =
                rightToLeft;

            if (tabStops != null)
            {
                for (int i = 0;
                     i < tabStops.Count;
                     i++)
                {
                    line.TabStops.Add(
                        tabStops[i]);
                }
            }

            return line;
        }

        private static float CalculateRichTabAdvance(
            RichTextLine line,
            float currentOffset)
        {
            if (line == null)
                return 1f;

            for (int i = 0;
                 i < line.TabStops.Count;
                 i++)
            {
                RichTabStop stop =
                    line.TabStops[i];

                if (stop != null &&
                    stop.Position >
                    currentOffset + 0.5f)
                {
                    return Math.Max(
                        1f,
                        stop.Position -
                        currentOffset);
                }
            }

            float tabSize =
                Math.Max(
                    8f,
                    line.DefaultTabSize);

            float next =
                ((float)Math.Floor(
                    currentOffset /
                    tabSize) +
                 1f) *
                tabSize;

            return Math.Max(
                1f,
                next -
                currentOffset);
        }

        private static void AddRichBulletToken(
            List<RichTextToken> tokens,
            XmlNode pPr,
            RichRunStyle baseStyle,
            int paragraphIndex)
        {
            if (pPr == null || DirectChild(pPr, "buNone") != null)
                return;

            XmlNode bulletChar = DirectChild(pPr, "buChar");
            if (bulletChar != null)
            {
                string value = GetAttr(bulletChar, "char");
                if (string.IsNullOrEmpty(value))
                    value = "•";

                AddRichTextTokens(tokens, value + " ", baseStyle.Clone());
                return;
            }

            XmlNode auto = DirectChild(pPr, "buAutoNum");
            if (auto != null)
            {
                int start = (int)GetLong(auto, "startAt", 1);
                string type = GetAttr(auto, "type");
                string number = FormatRichAutoNumber(type, start + paragraphIndex);
                AddRichTextTokens(tokens, number + " ", baseStyle.Clone());
            }
        }

        private static string FormatRichAutoNumber(string type, int value)
        {
            value = Math.Max(1, value);

            if (!string.IsNullOrEmpty(type) &&
                type.IndexOf("alpha", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                int n = (value - 1) % 26;
                char c = (char)('a' + n);
                if (type.IndexOf("Uc", StringComparison.OrdinalIgnoreCase) >= 0)
                    c = char.ToUpperInvariant(c);
                return type.IndexOf("Paren", StringComparison.OrdinalIgnoreCase) >= 0
                    ? c.ToString() + ")"
                    : c.ToString() + ".";
            }

            if (!string.IsNullOrEmpty(type) &&
                type.IndexOf("roman", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                string roman = ToRichRoman(value);
                if (type.IndexOf("Uc", StringComparison.OrdinalIgnoreCase) < 0)
                    roman = roman.ToLowerInvariant();
                return type.IndexOf("Paren", StringComparison.OrdinalIgnoreCase) >= 0
                    ? roman + ")"
                    : roman + ".";
            }

            return type != null && type.IndexOf("Paren", StringComparison.OrdinalIgnoreCase) >= 0
                ? value.ToString() + ")"
                : value.ToString() + ".";
        }

        private static string ToRichRoman(int value)
        {
            value = Math.Max(1, Math.Min(3999, value));
            int[] values = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] numerals = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            StringBuilder result = new StringBuilder();

            for (int i = 0; i < values.Length; i++)
            {
                while (value >= values[i])
                {
                    result.Append(numerals[i]);
                    value -= values[i];
                }
            }

            return result.ToString();
        }

        private static void AddRichTextTokens(
            List<RichTextToken> tokens,
            string text,
            RichRunStyle style)
        {
            if (string.IsNullOrEmpty(text))
                return;

            StringBuilder buffer = new StringBuilder();
            bool bufferWhitespace = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (c == '\r')
                    continue;

                if (c == '\n')
                {
                    FlushRichToken(tokens, buffer, style);
                    tokens.Add(new RichTextToken { Break = true });
                    continue;
                }

                bool whitespace = char.IsWhiteSpace(c);
                bool cjk = IsRichCjkCharacter(c);

                if (cjk)
                {
                    FlushRichToken(tokens, buffer, style);
                    tokens.Add(new RichTextToken
                    {
                        Text = c.ToString(),
                        Style = style.Clone()
                    });
                    continue;
                }

                if (buffer.Length > 0 && whitespace != bufferWhitespace)
                    FlushRichToken(tokens, buffer, style);

                bufferWhitespace = whitespace;
                buffer.Append(c);
            }

            FlushRichToken(tokens, buffer, style);
        }

        private static void FlushRichToken(
            List<RichTextToken> tokens,
            StringBuilder buffer,
            RichRunStyle style)
        {
            if (buffer.Length == 0)
                return;

            tokens.Add(new RichTextToken
            {
                Text = buffer.ToString(),
                Style = style.Clone()
            });
            buffer.Length = 0;
        }

        private static bool IsRichCjkCharacter(char value)
        {
            int c = value;
            return (c >= 0xAC00 && c <= 0xD7AF) ||
                (c >= 0x3130 && c <= 0x318F) ||
                (c >= 0x4E00 && c <= 0x9FFF) ||
                (c >= 0x3040 && c <= 0x30FF);
        }

        private static RichRunStyle ApplyRichRunProperties(
            RichRunStyle source,
            XmlNode rPr,
            Dictionary<string, Color> theme,
            string text,
            float fontScale)
        {
            RichRunStyle result = source != null ? source.Clone() : new RichRunStyle();

            if (!string.IsNullOrEmpty(text) &&
                ContainsCjk(text) &&
                string.Equals(result.FontName, "Arial", StringComparison.OrdinalIgnoreCase))
            {
                result.FontName = "Malgun Gothic";
            }

            if (rPr != null)
            {
                long size = GetLong(rPr, "sz", 0);
                if (size > 0)
                    result.SizePt = (float)(size / 100.0);

                string bold = GetAttr(rPr, "b");
                string italic = GetAttr(rPr, "i");
                string underline = GetAttr(rPr, "u");

                if (!string.IsNullOrEmpty(bold))
                {
                    if (IsRichBooleanTrue(bold))
                        result.FontStyle |= FontStyle.Bold;
                    else
                        result.FontStyle &= ~FontStyle.Bold;
                }

                if (!string.IsNullOrEmpty(italic))
                {
                    if (IsRichBooleanTrue(italic))
                        result.FontStyle |= FontStyle.Italic;
                    else
                        result.FontStyle &= ~FontStyle.Italic;
                }

                if (!string.IsNullOrEmpty(underline))
                {
                    if (!string.Equals(underline, "none", StringComparison.OrdinalIgnoreCase))
                        result.FontStyle |= FontStyle.Underline;
                    else
                        result.FontStyle &= ~FontStyle.Underline;
                }

                XmlNode typefaceNode = ContainsCjk(text)
                    ? DirectChild(rPr, "ea")
                    : DirectChild(rPr, "latin");

                if (typefaceNode == null)
                    typefaceNode = DirectChild(rPr, "latin");
                if (typefaceNode == null)
                    typefaceNode = DirectChild(rPr, "ea");

                if (typefaceNode != null)
                {
                    string typeface = GetAttr(typefaceNode, "typeface");
                    string resolved = ResolveRichThemeFont(typeface);
                    if (!string.IsNullOrWhiteSpace(resolved))
                        result.FontName = resolved;
                }

                Color? color = ReadSolidFill(rPr, theme);
                if (color.HasValue)
                    result.Color = color.Value;
            }

            result.SizePt = Math.Max(4f, Math.Min(200f, result.SizePt * fontScale));
            return result;
        }

        private static string ResolveRichThemeFont(string typeface)
        {
            if (string.IsNullOrWhiteSpace(typeface))
                return null;

            if (!typeface.StartsWith("+", StringComparison.Ordinal))
                return typeface;

            if (typeface.StartsWith("+mj-ea", StringComparison.OrdinalIgnoreCase) &&
                ThemeFonts.ContainsKey("majorEA"))
                return ThemeFonts["majorEA"];

            if (typeface.StartsWith("+mn-ea", StringComparison.OrdinalIgnoreCase) &&
                ThemeFonts.ContainsKey("minorEA"))
                return ThemeFonts["minorEA"];

            if (typeface.StartsWith("+mj", StringComparison.OrdinalIgnoreCase) &&
                ThemeFonts.ContainsKey("major"))
                return ThemeFonts["major"];

            if (typeface.StartsWith("+mn", StringComparison.OrdinalIgnoreCase) &&
                ThemeFonts.ContainsKey("minor"))
                return ThemeFonts["minor"];

            return null;
        }

        private static bool IsRichBooleanTrue(string value)
        {
            return value == "1" ||
                string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        private static void MeasureRichToken(
            Graphics g,
            RichTextToken token,
            out float width,
            out float height)
        {
            width = 0f;
            height = 1f;

            if (token == null || token.Style == null || string.IsNullOrEmpty(token.Text))
                return;

            using (Font font = SafeFont(token.Style.FontName, token.Style.SizePt, token.Style.FontStyle))
            using (StringFormat format = (StringFormat)StringFormat.GenericTypographic.Clone())
            {
                format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
                SizeF measured = g.MeasureString(token.Text, font, int.MaxValue, format);
                width = Math.Max(0f, measured.Width);
                height = Math.Max(font.Height, measured.Height);
            }
        }

        private static float RichFontHeight(Graphics g, RichRunStyle style)
        {
            if (style == null)
                return 1f;

            using (Font font = SafeFont(style.FontName, style.SizePt, style.FontStyle))
                return Math.Max(1f, font.Height);
        }

        private static void TrimTrailingRichWhitespace(Graphics g, RichTextLine line)
        {
            if (line == null)
                return;

            while (line.Tokens.Count > 0)
            {
                RichTextToken last = line.Tokens[line.Tokens.Count - 1];
                if (last == null || string.IsNullOrEmpty(last.Text) || !string.IsNullOrWhiteSpace(last.Text))
                    break;

                line.Tokens.RemoveAt(line.Tokens.Count - 1);
            }

            line.Width = 0f;
            float measuredHeight = 1f;

            for (int i = 0; i < line.Tokens.Count; i++)
            {
                float width;
                float height;

                if (line.Tokens[i] != null &&
                    line.Tokens[i].Tab)
                {
                    width =
                        CalculateRichTabAdvance(
                            line,
                            line.LeftOffset +
                            line.Width);
                    height =
                        RichFontHeight(
                            g,
                            line.Tokens[i].Style);
                }
                else
                {
                    MeasureRichToken(
                        g,
                        line.Tokens[i],
                        out width,
                        out height);
                }

                line.Width += width;
                measuredHeight = Math.Max(measuredHeight, height);
            }

            line.Height = Math.Max(line.Height, measuredHeight);
        }

        private static float CalculateRichAlignedTabAdvance(
            Graphics g,
            RichTextLine line,
            int tabTokenIndex,
            float currentOffset)
        {
            if (line == null)
                return 1f;

            RichTabStop selected =
                null;

            for (int i = 0;
                 i < line.TabStops.Count;
                 i++)
            {
                RichTabStop stop =
                    line.TabStops[i];

                if (stop != null &&
                    stop.Position >
                    currentOffset + 0.5f)
                {
                    selected =
                        stop;
                    break;
                }
            }

            if (selected == null)
            {
                return CalculateRichTabAdvance(
                    line,
                    currentOffset);
            }

            string alignment =
                selected.Alignment ??
                "l";

            if (alignment == "l")
            {
                return Math.Max(
                    1f,
                    selected.Position -
                    currentOffset);
            }

            float segmentWidth;
            float decimalPrefixWidth;
            bool hasDecimal;

            MeasureRichTabSegment(
                g,
                line,
                tabTokenIndex + 1,
                out segmentWidth,
                out decimalPrefixWidth,
                out hasDecimal);

            float desiredStart =
                selected.Position;

            if (alignment == "ctr")
            {
                desiredStart -=
                    segmentWidth /
                    2f;
            }
            else if (alignment == "r")
            {
                desiredStart -=
                    segmentWidth;
            }
            else if (alignment == "dec")
            {
                desiredStart -=
                    hasDecimal
                        ? decimalPrefixWidth
                        : segmentWidth;
            }

            return Math.Max(
                1f,
                desiredStart -
                currentOffset);
        }

        private static void MeasureRichTabSegment(
            Graphics g,
            RichTextLine line,
            int startIndex,
            out float totalWidth,
            out float decimalPrefixWidth,
            out bool hasDecimal)
        {
            totalWidth = 0f;
            decimalPrefixWidth = 0f;
            hasDecimal = false;

            if (line == null)
                return;

            for (int i = startIndex;
                 i < line.Tokens.Count;
                 i++)
            {
                RichTextToken token =
                    line.Tokens[i];

                if (token == null ||
                    token.Tab ||
                    token.Break)
                {
                    break;
                }

                if (token.Style == null ||
                    string.IsNullOrEmpty(
                        token.Text))
                {
                    continue;
                }

                if (!hasDecimal)
                {
                    int decimalIndex =
                        token.Text.IndexOf('.');

                    if (decimalIndex >= 0)
                    {
                        string prefix =
                            token.Text.Substring(
                                0,
                                decimalIndex);

                        RichTextToken prefixToken =
                            new RichTextToken();

                        prefixToken.Text =
                            prefix;
                        prefixToken.Style =
                            token.Style;

                        float prefixWidth;
                        float prefixHeight;

                        MeasureRichToken(
                            g,
                            prefixToken,
                            out prefixWidth,
                            out prefixHeight);

                        decimalPrefixWidth =
                            totalWidth +
                            prefixWidth;

                        hasDecimal =
                            true;
                    }
                }

                float width;
                float height;

                MeasureRichToken(
                    g,
                    token,
                    out width,
                    out height);

                totalWidth +=
                    width;
            }

            if (!hasDecimal)
            {
                decimalPrefixWidth =
                    totalWidth;
            }
        }

        private static void DrawRichTextLine(
            Graphics g,
            RichTextLine line,
            float startX,
            float y)
        {
            if (line == null)
                return;

            float consumed = 0f;
            float x = line.RightToLeft
                ? startX + line.Width
                : startX;

            for (int i = 0;
                 i < line.Tokens.Count;
                 i++)
            {
                RichTextToken token =
                    line.Tokens[i];

                if (token == null)
                    continue;

                if (token.Tab)
                {
                    float advance =
                        CalculateRichAlignedTabAdvance(
                            g,
                            line,
                            i,
                            line.LeftOffset +
                            consumed);

                    consumed += advance;

                    if (line.RightToLeft)
                        x -= advance;
                    else
                        x += advance;

                    continue;
                }

                if (token.Style == null ||
                    string.IsNullOrEmpty(
                        token.Text))
                {
                    continue;
                }

                float width;
                float height;
                MeasureRichToken(
                    g,
                    token,
                    out width,
                    out height);

                float drawX = x;

                if (line.RightToLeft)
                {
                    x -= width;
                    drawX = x;
                }

                using (Font font = SafeFont(
                    token.Style.FontName,
                    token.Style.SizePt,
                    token.Style.FontStyle))
                using (Brush brush =
                    new SolidBrush(
                        token.Style.Color))
                using (StringFormat format =
                    (StringFormat)
                    StringFormat.GenericTypographic.Clone())
                {
                    format.FormatFlags |=
                        StringFormatFlags.MeasureTrailingSpaces;

                    if (line.RightToLeft)
                    {
                        format.FormatFlags |=
                            StringFormatFlags.DirectionRightToLeft;
                    }

                    g.DrawString(
                        token.Text,
                        font,
                        brush,
                        drawX,
                        y,
                        format);
                }

                if (!line.RightToLeft)
                    x += width;

                consumed += width;
            }
        }

        private static float ReadRichSpacingPixels(
            Graphics g,
            XmlNode spacingNode,
            float referenceSizePt)
        {
            if (spacingNode == null)
                return 0f;

            XmlNode points = DirectChild(spacingNode, "spcPts");
            if (points != null)
            {
                long raw = GetLong(points, "val", 0);
                return RichPointsToPixels(g, raw / 100f);
            }

            XmlNode percent = DirectChild(spacingNode, "spcPct");
            if (percent != null)
            {
                long raw = GetLong(percent, "val", 0);
                float ratio = Math.Max(0f, raw / 100000f);
                return RichPointsToPixels(g, referenceSizePt * ratio);
            }

            return 0f;
        }

        private static void ReadRichLineSpacing(
            Graphics g,
            XmlNode spacingNode,
            float referenceSizePt,
            float lineReduction,
            RichParagraphLayout result)
        {
            result.LineFactor = Math.Max(0.2f, 1f - lineReduction);
            result.AbsoluteLineHeight = 0f;

            if (spacingNode == null)
                return;

            XmlNode points = DirectChild(spacingNode, "spcPts");
            if (points != null)
            {
                long raw = GetLong(points, "val", 0);
                result.AbsoluteLineHeight = RichPointsToPixels(g, raw / 100f);
                return;
            }

            XmlNode percent = DirectChild(spacingNode, "spcPct");
            if (percent != null)
            {
                long raw = GetLong(percent, "val", 100000);
                result.LineFactor = Math.Max(0.2f, raw / 100000f);
                result.LineFactor *= Math.Max(0.2f, 1f - lineReduction);
            }
        }

        private static float RichPointsToPixels(Graphics g, float points)
        {
            float dpi = g != null ? g.DpiY : 96f;
            return Math.Max(0f, points * dpi / 72f);
        }
    }
}
