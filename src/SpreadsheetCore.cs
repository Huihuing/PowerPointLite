using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;

namespace PptxViewer
{
    internal enum SpreadsheetCellKind
    {
        Blank,
        Text,
        Number,
        Boolean,
        Formula
    }

    internal sealed class SpreadsheetCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public SpreadsheetCellKind Kind { get; set; }
        public string Text { get; set; }
        public double NumberValue { get; set; }
        public bool BooleanValue { get; set; }
        public string Formula { get; set; }

        public SpreadsheetCell()
        {
            Row = 1;
            Column = 1;
            Kind = SpreadsheetCellKind.Blank;
            Text = string.Empty;
            Formula = string.Empty;
        }

        public string DisplayText
        {
            get
            {
                if (Kind == SpreadsheetCellKind.Text)
                    return Text ?? string.Empty;
                if (Kind == SpreadsheetCellKind.Number)
                    return NumberValue.ToString("G15", CultureInfo.InvariantCulture);
                if (Kind == SpreadsheetCellKind.Boolean)
                    return BooleanValue ? "TRUE" : "FALSE";
                if (Kind == SpreadsheetCellKind.Formula)
                    return "=" + (Formula ?? string.Empty);
                return string.Empty;
            }
        }

        public SpreadsheetCell Clone()
        {
            SpreadsheetCell copy = new SpreadsheetCell();
            copy.Row = Row;
            copy.Column = Column;
            copy.Kind = Kind;
            copy.Text = Text;
            copy.NumberValue = NumberValue;
            copy.BooleanValue = BooleanValue;
            copy.Formula = Formula;
            return copy;
        }
    }

    internal sealed class SpreadsheetSheet
    {
        private readonly Dictionary<long, SpreadsheetCell> cells =
            new Dictionary<long, SpreadsheetCell>();

        public string Name { get; set; }

        public SpreadsheetSheet()
        {
            Name = "Sheet";
        }

        private static long CellKey(int row, int column)
        {
            return ((long)row << 32) | (uint)column;
        }

        public SpreadsheetCell GetCell(int row, int column, bool create)
        {
            if (row < 1 || column < 1)
                return null;

            long key = CellKey(row, column);
            SpreadsheetCell cell;

            if (cells.TryGetValue(key, out cell))
                return cell;

            if (!create)
                return null;

            cell = new SpreadsheetCell();
            cell.Row = row;
            cell.Column = column;
            cells[key] = cell;
            return cell;
        }

        public void SetInput(int row, int column, string input)
        {
            SpreadsheetCell cell = GetCell(row, column, true);
            string value = input ?? string.Empty;

            if (string.IsNullOrEmpty(value))
            {
                cells.Remove(CellKey(row, column));
                return;
            }

            if (value.StartsWith("=", StringComparison.Ordinal) && value.Length > 1)
            {
                cell.Kind = SpreadsheetCellKind.Formula;
                cell.Formula = value.Substring(1);
                cell.Text = string.Empty;
                return;
            }

            bool booleanValue;
            if (bool.TryParse(value, out booleanValue))
            {
                cell.Kind = SpreadsheetCellKind.Boolean;
                cell.BooleanValue = booleanValue;
                cell.Text = string.Empty;
                cell.Formula = string.Empty;
                return;
            }

            double number;
            if (double.TryParse(
                    value,
                    NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.InvariantCulture,
                    out number))
            {
                cell.Kind = SpreadsheetCellKind.Number;
                cell.NumberValue = number;
                cell.Text = string.Empty;
                cell.Formula = string.Empty;
                return;
            }

            cell.Kind = SpreadsheetCellKind.Text;
            cell.Text = value;
            cell.Formula = string.Empty;
        }

        public List<SpreadsheetCell> GetCellsSorted()
        {
            List<SpreadsheetCell> result =
                new List<SpreadsheetCell>(cells.Values);

            result.Sort(delegate(SpreadsheetCell left, SpreadsheetCell right)
            {
                int rowCompare = left.Row.CompareTo(right.Row);
                return rowCompare != 0
                    ? rowCompare
                    : left.Column.CompareTo(right.Column);
            });

            return result;
        }

        public int MaxRow
        {
            get
            {
                int value = 1;
                foreach (SpreadsheetCell cell in cells.Values)
                    value = Math.Max(value, cell.Row);
                return value;
            }
        }

        public int MaxColumn
        {
            get
            {
                int value = 1;
                foreach (SpreadsheetCell cell in cells.Values)
                    value = Math.Max(value, cell.Column);
                return value;
            }
        }

        public SpreadsheetSheet Clone()
        {
            SpreadsheetSheet copy = new SpreadsheetSheet();
            copy.Name = Name;

            foreach (SpreadsheetCell cell in cells.Values)
            {
                SpreadsheetCell cloned = cell.Clone();
                copy.cells[CellKey(cloned.Row, cloned.Column)] = cloned;
            }

            return copy;
        }
    }

    internal sealed class SpreadsheetDocument
    {
        private readonly List<SpreadsheetSheet> sheets =
            new List<SpreadsheetSheet>();

        public string Title { get; set; }

        public IList<SpreadsheetSheet> Sheets
        {
            get { return sheets; }
        }

        public SpreadsheetDocument()
        {
            Title = "New Workbook";
        }

        public static SpreadsheetDocument CreateNew(string title)
        {
            SpreadsheetDocument document = new SpreadsheetDocument();
            if (!string.IsNullOrEmpty(title))
                document.Title = title;
            document.AddSheet("Sheet1");
            return document;
        }

        public SpreadsheetSheet AddSheet(string name)
        {
            SpreadsheetSheet sheet = new SpreadsheetSheet();
            sheet.Name = MakeUniqueSheetName(name);
            sheets.Add(sheet);
            return sheet;
        }

        public bool RemoveSheet(int index)
        {
            if (index < 0 || index >= sheets.Count || sheets.Count <= 1)
                return false;
            sheets.RemoveAt(index);
            return true;
        }

        public bool MoveSheet(int fromIndex, int toIndex)
        {
            if (fromIndex < 0 || fromIndex >= sheets.Count ||
                toIndex < 0 || toIndex >= sheets.Count ||
                fromIndex == toIndex)
            {
                return false;
            }

            SpreadsheetSheet sheet = sheets[fromIndex];
            sheets.RemoveAt(fromIndex);
            sheets.Insert(toIndex, sheet);
            return true;
        }

        public bool RenameSheet(int index, string name)
        {
            if (index < 0 || index >= sheets.Count)
                return false;

            string candidate = SanitizeSheetName(name);
            if (string.IsNullOrEmpty(candidate))
                return false;

            for (int i = 0; i < sheets.Count; i++)
            {
                if (i != index &&
                    string.Equals(sheets[i].Name, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            sheets[index].Name = candidate;
            return true;
        }

        private string MakeUniqueSheetName(string value)
        {
            string baseName = SanitizeSheetName(value);
            if (string.IsNullOrEmpty(baseName))
                baseName = "Sheet";

            string candidate = baseName;
            int suffix = 2;

            while (ContainsSheetName(candidate))
            {
                string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                int maxBase = Math.Max(1, 31 - suffixText.Length - 1);
                string trimmed = baseName.Length > maxBase
                    ? baseName.Substring(0, maxBase)
                    : baseName;
                candidate = trimmed + " " + suffixText;
                suffix++;
            }

            return candidate;
        }

        private bool ContainsSheetName(string value)
        {
            for (int i = 0; i < sheets.Count; i++)
            {
                if (string.Equals(sheets[i].Name, value, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        public static string SanitizeSheetName(string value)
        {
            string text = (value ?? string.Empty).Trim();
            char[] invalid = new char[] { ':', '\\', '/', '?', '*', '[', ']' };

            for (int i = 0; i < invalid.Length; i++)
                text = text.Replace(invalid[i].ToString(), string.Empty);

            text = text.Trim('\'');
            if (text.Length > 31)
                text = text.Substring(0, 31);
            return text;
        }

        public SpreadsheetDocument Clone()
        {
            SpreadsheetDocument copy = new SpreadsheetDocument();
            copy.Title = Title;
            for (int i = 0; i < sheets.Count; i++)
                copy.sheets.Add(sheets[i].Clone());
            return copy;
        }
    }

    internal sealed class XlsxEditSafetyResult
    {
        public bool IsProjectGenerated { get; set; }
        public bool HasUnsupportedContent { get; set; }
        public bool CanEditSafely { get; set; }
        public string Warning { get; set; }

        public XlsxEditSafetyResult()
        {
            Warning = string.Empty;
        }
    }

    internal static class XlsxEditSafety
    {
        public static XlsxEditSafetyResult Analyze(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("XLSX file was not found.", path);

            XlsxEditSafetyResult result = new XlsxEditSafetyResult();

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                result.IsProjectGenerated = IsProjectGenerated(archive);

                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    string part = OpcPackageUtility.NormalizePartName(
                        archive.Entries[i].FullName);

                    if (IsSupportedPart(part))
                        continue;

                    if (part.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                        result.HasUnsupportedContent = true;
                }
            }

            result.CanEditSafely =
                result.IsProjectGenerated &&
                !result.HasUnsupportedContent;

            if (!result.CanEditSafely)
            {
                result.Warning = result.IsProjectGenerated
                    ? "This workbook contains XLSX parts not yet preserved by the experimental editor. It will remain unchanged."
                    : "This workbook was not generated by the current project writer. Open/edit support stays disabled until unknown-part preservation is implemented.";
            }

            return result;
        }

        private static bool IsProjectGenerated(ZipArchive archive)
        {
            XmlDocument core = OpcPackageUtility.ReadXmlPart(archive, "docProps/core.xml");
            if (core == null || core.DocumentElement == null)
                return false;

            XmlNode creator = FindFirst(core.DocumentElement, "creator");
            XmlNode modifiedBy = FindFirst(core.DocumentElement, "lastModifiedBy");

            return ContainsProjectName(creator) || ContainsProjectName(modifiedBy);
        }

        private static bool ContainsProjectName(XmlNode node)
        {
            string text = node == null ? string.Empty : node.InnerText;
            return text.IndexOf(
                "PowerPointLite",
                StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsSupportedPart(string part)
        {
            if (string.IsNullOrEmpty(part))
                return true;

            if (part == "[Content_Types].xml" ||
                part == "_rels/.rels" ||
                part == "docProps/core.xml" ||
                part == "docProps/app.xml" ||
                part == "xl/workbook.xml" ||
                part == "xl/_rels/workbook.xml.rels" ||
                part == "xl/styles.xml" ||
                part == "xl/sharedStrings.xml")
            {
                return true;
            }

            return part.StartsWith(
                "xl/worksheets/sheet",
                StringComparison.OrdinalIgnoreCase) &&
                part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
        }

        private static XmlNode FindFirst(XmlNode node, string localName)
        {
            if (node == null)
                return null;
            if (node.LocalName == localName)
                return node;

            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                XmlNode found = FindFirst(node.ChildNodes[i], localName);
                if (found != null)
                    return found;
            }
            return null;
        }
    }

    internal static class XlsxWriter
    {
        private const string MainNs =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string RelNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string OfficeDocumentRel =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string WorksheetRel =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        private const string StylesRel =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

        public static void Save(SpreadsheetDocument document, string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path is required.", "outputPath");
            if (document.Sheets.Count == 0)
                throw new InvalidOperationException("A workbook must contain at least one worksheet.");

            string fullPath = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string stage = fullPath + ".writing";
            if (File.Exists(stage))
                File.Delete(stage);

            try
            {
                using (ZipArchive archive = ZipFile.Open(stage, ZipArchiveMode.Create))
                {
                    WriteContentTypes(archive, document.Sheets.Count);
                    WritePackageRelationships(archive);
                    WriteCoreProperties(archive, document.Title);
                    WriteAppProperties(archive, document.Sheets.Count);
                    WriteWorkbook(archive, document);
                    WriteWorkbookRelationships(archive, document.Sheets.Count);
                    WriteStyles(archive);

                    for (int i = 0; i < document.Sheets.Count; i++)
                        WriteWorksheet(archive, i + 1, document.Sheets[i]);
                }

                ReplaceSafely(stage, fullPath);
            }
            catch
            {
                try { if (File.Exists(stage)) File.Delete(stage); }
                catch { }
                throw;
            }
        }

        private static void WriteContentTypes(ZipArchive archive, int sheetCount)
        {
            Dictionary<string, string> defaults = new Dictionary<string, string>();
            defaults.Add("rels", "application/vnd.openxmlformats-package.relationships+xml");
            defaults.Add("xml", "application/xml");

            Dictionary<string, string> overrides = new Dictionary<string, string>();
            overrides.Add(
                "xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            overrides.Add(
                "xl/styles.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            overrides.Add(
                "docProps/core.xml",
                "application/vnd.openxmlformats-package.core-properties+xml");
            overrides.Add(
                "docProps/app.xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml");

            for (int i = 1; i <= sheetCount; i++)
            {
                overrides.Add(
                    "xl/worksheets/sheet" + i.ToString() + ".xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            }

            OpcPackageUtility.WriteXmlPart(
                archive,
                "[Content_Types].xml",
                OpcPackageUtility.CreateContentTypesDocument(defaults, overrides));
        }

        private static void WritePackageRelationships(ZipArchive archive)
        {
            List<OpcRelationship> rels = new List<OpcRelationship>();
            rels.Add(MakeRelationship("rId1", OfficeDocumentRel, "xl/workbook.xml"));
            rels.Add(MakeRelationship(
                "rId2",
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                "docProps/core.xml"));
            rels.Add(MakeRelationship(
                "rId3",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                "docProps/app.xml"));
            OpcPackageUtility.WriteXmlPart(
                archive,
                "_rels/.rels",
                OpcPackageUtility.CreateRelationshipsDocument(rels));
        }

        private static void WriteWorkbook(ZipArchive archive, SpreadsheetDocument document)
        {
            StringBuilder sheets = new StringBuilder();
            for (int i = 0; i < document.Sheets.Count; i++)
            {
                SpreadsheetSheet sheet = document.Sheets[i];
                sheets.Append("<sheet name=\"");
                sheets.Append(EscapeXml(sheet == null ? "Sheet" + (i + 1).ToString() : sheet.Name));
                sheets.Append("\" sheetId=\"");
                sheets.Append((i + 1).ToString());
                sheets.Append("\" r:id=\"rId");
                sheets.Append((i + 1).ToString());
                sheets.Append("\"/>");
            }

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<workbook xmlns=\"" + MainNs + "\" xmlns:r=\"" + RelNs + "\">" +
                "<bookViews><workbookView/></bookViews>" +
                "<sheets>" + sheets.ToString() + "</sheets>" +
                "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>" +
                "</workbook>";
            WriteXmlString(archive, "xl/workbook.xml", xml);
        }

        private static void WriteWorkbookRelationships(ZipArchive archive, int sheetCount)
        {
            List<OpcRelationship> rels = new List<OpcRelationship>();

            for (int i = 1; i <= sheetCount; i++)
            {
                rels.Add(MakeRelationship(
                    "rId" + i.ToString(),
                    WorksheetRel,
                    "worksheets/sheet" + i.ToString() + ".xml"));
            }

            rels.Add(MakeRelationship(
                "rId" + (sheetCount + 1).ToString(),
                StylesRel,
                "styles.xml"));

            OpcPackageUtility.WriteXmlPart(
                archive,
                "xl/_rels/workbook.xml.rels",
                OpcPackageUtility.CreateRelationshipsDocument(rels));
        }

        private static void WriteStyles(ZipArchive archive)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<styleSheet xmlns=\"" + MainNs + "\">" +
                "<numFmts count=\"0\"/>" +
                "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Arial\"/><family val=\"2\"/></font></fonts>" +
                "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                "</styleSheet>";
            WriteXmlString(archive, "xl/styles.xml", xml);
        }

        private static void WriteWorksheet(
            ZipArchive archive,
            int sheetNumber,
            SpreadsheetSheet sheet)
        {
            if (sheet == null)
                sheet = new SpreadsheetSheet();

            List<SpreadsheetCell> cells = sheet.GetCellsSorted();
            StringBuilder body = new StringBuilder();
            int currentRow = -1;

            for (int i = 0; i < cells.Count; i++)
            {
                SpreadsheetCell cell = cells[i];
                if (cell == null || cell.Kind == SpreadsheetCellKind.Blank)
                    continue;

                if (cell.Row != currentRow)
                {
                    if (currentRow >= 0)
                        body.Append("</row>");
                    currentRow = cell.Row;
                    body.Append("<row r=\"");
                    body.Append(currentRow.ToString());
                    body.Append("\">");
                }

                body.Append(BuildCellXml(cell));
            }

            if (currentRow >= 0)
                body.Append("</row>");

            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"" + MainNs + "\">" +
                "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>" +
                "<sheetFormatPr defaultRowHeight=\"15\"/>" +
                "<sheetData>" + body.ToString() + "</sheetData>" +
                "</worksheet>";

            WriteXmlString(
                archive,
                "xl/worksheets/sheet" + sheetNumber.ToString() + ".xml",
                xml);
        }

        private static string BuildCellXml(SpreadsheetCell cell)
        {
            string reference = CellReference(cell.Row, cell.Column);

            if (cell.Kind == SpreadsheetCellKind.Text)
            {
                return "<c r=\"" + reference + "\" t=\"inlineStr\"><is><t xml:space=\"preserve\">" +
                    EscapeXml(cell.Text) + "</t></is></c>";
            }

            if (cell.Kind == SpreadsheetCellKind.Boolean)
            {
                return "<c r=\"" + reference + "\" t=\"b\"><v>" +
                    (cell.BooleanValue ? "1" : "0") + "</v></c>";
            }

            if (cell.Kind == SpreadsheetCellKind.Formula)
            {
                return "<c r=\"" + reference + "\"><f>" +
                    EscapeXml(cell.Formula) + "</f><v>0</v></c>";
            }

            return "<c r=\"" + reference + "\"><v>" +
                cell.NumberValue.ToString("R", CultureInfo.InvariantCulture) +
                "</v></c>";
        }

        private static void WriteCoreProperties(ZipArchive archive, string title)
        {
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                "<dc:title>" + EscapeXml(title) + "</dc:title>" +
                "<dc:creator>PowerPointLite</dc:creator><cp:lastModifiedBy>PowerPointLite</cp:lastModifiedBy>" +
                "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + now + "</dcterms:created>" +
                "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + now + "</dcterms:modified>" +
                "</cp:coreProperties>";
            WriteXmlString(archive, "docProps/core.xml", xml);
        }

        private static void WriteAppProperties(ZipArchive archive, int sheetCount)
        {
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
                "<Application>PowerPointLite</Application><DocSecurity>0</DocSecurity>" +
                "<AppVersion>1.0</AppVersion><Company></Company>" +
                "</Properties>";
            WriteXmlString(archive, "docProps/app.xml", xml);
        }

        private static void ReplaceSafely(string stage, string destination)
        {
            if (!File.Exists(destination))
            {
                File.Move(stage, destination);
                return;
            }

            string backup = destination + ".backup";
            if (File.Exists(backup))
                File.Delete(backup);
            File.Move(destination, backup);

            try
            {
                File.Move(stage, destination);
                File.Delete(backup);
            }
            catch
            {
                try
                {
                    if (File.Exists(destination))
                        File.Delete(destination);
                    if (File.Exists(backup))
                        File.Move(backup, destination);
                }
                catch { }
                throw;
            }
        }

        internal static string CellReference(int row, int column)
        {
            int value = Math.Max(1, column);
            StringBuilder letters = new StringBuilder();

            while (value > 0)
            {
                value--;
                letters.Insert(0, (char)('A' + (value % 26)));
                value /= 26;
            }

            return letters.ToString() + Math.Max(1, row).ToString();
        }

        private static OpcRelationship MakeRelationship(
            string id,
            string type,
            string target)
        {
            OpcRelationship relationship = new OpcRelationship();
            relationship.Id = id;
            relationship.Type = type;
            relationship.Target = target;
            return relationship;
        }

        private static void WriteXmlString(ZipArchive archive, string part, string xml)
        {
            XmlDocument document = new XmlDocument();
            document.PreserveWhitespace = true;
            document.LoadXml(xml);
            OpcPackageUtility.WriteXmlPart(archive, part, document);
        }

        private static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;
        }
    }

    internal static class XlsxReader
    {
        private const string WorksheetRelSuffix = "/worksheet";

        public static SpreadsheetDocument Read(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("XLSX file was not found.", path);

            SpreadsheetDocument document = new SpreadsheetDocument();

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                document.Title = ReadCoreTitle(archive);
                List<string> sharedStrings = ReadSharedStrings(archive);
                XmlDocument workbook = OpcPackageUtility.ReadXmlPart(archive, "xl/workbook.xml");
                if (workbook == null)
                    throw new InvalidDataException("xl/workbook.xml is missing or invalid.");

                List<OpcRelationship> rels =
                    OpcPackageUtility.ReadRelationships(archive, "xl/workbook.xml");
                Dictionary<string, OpcRelationship> byId =
                    new Dictionary<string, OpcRelationship>(StringComparer.Ordinal);

                for (int i = 0; i < rels.Count; i++)
                {
                    if (!string.IsNullOrEmpty(rels[i].Id))
                        byId[rels[i].Id] = rels[i];
                }

                XmlNode sheets = FindFirst(workbook.DocumentElement, "sheets");
                if (sheets == null)
                    throw new InvalidDataException("Workbook sheet list is missing.");

                for (int i = 0; i < sheets.ChildNodes.Count; i++)
                {
                    XmlNode sheetNode = sheets.ChildNodes[i];
                    if (sheetNode.LocalName != "sheet")
                        continue;

                    string name = GetAttribute(sheetNode, "name");
                    string relId = GetRelationshipId(sheetNode);
                    OpcRelationship relationship;
                    if (string.IsNullOrEmpty(relId) ||
                        !byId.TryGetValue(relId, out relationship) ||
                        relationship == null ||
                        relationship.IsExternal ||
                        !relationship.Type.EndsWith(WorksheetRelSuffix, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    SpreadsheetSheet sheet = new SpreadsheetSheet();
                    sheet.Name = string.IsNullOrEmpty(name)
                        ? "Sheet" + (document.Sheets.Count + 1).ToString()
                        : name;
                    ReadWorksheet(
                        archive,
                        relationship.ResolvedPart,
                        sheet,
                        sharedStrings);
                    document.Sheets.Add(sheet);
                }
            }

            if (document.Sheets.Count == 0)
                document.AddSheet("Sheet1");

            return document;
        }

        private static void ReadWorksheet(
            ZipArchive archive,
            string part,
            SpreadsheetSheet target,
            List<string> sharedStrings)
        {
            XmlDocument worksheet = OpcPackageUtility.ReadXmlPart(archive, part);
            if (worksheet == null)
                return;

            List<XmlNode> cells = FindAll(worksheet.DocumentElement, "c");

            for (int i = 0; i < cells.Count; i++)
            {
                XmlNode node = cells[i];
                int row;
                int column;
                if (!TryParseReference(GetAttribute(node, "r"), out row, out column))
                    continue;

                string type = GetAttribute(node, "t");
                XmlNode formula = FindDirect(node, "f");
                XmlNode value = FindDirect(node, "v");
                SpreadsheetCell cell = target.GetCell(row, column, true);

                if (formula != null)
                {
                    cell.Kind = SpreadsheetCellKind.Formula;
                    cell.Formula = formula.InnerText ?? string.Empty;
                    continue;
                }

                if (type == "inlineStr")
                {
                    XmlNode text = FindFirst(FindDirect(node, "is"), "t");
                    cell.Kind = SpreadsheetCellKind.Text;
                    cell.Text = text == null ? string.Empty : text.InnerText;
                }
                else if (type == "s")
                {
                    int index;
                    cell.Kind = SpreadsheetCellKind.Text;
                    cell.Text = int.TryParse(value == null ? "" : value.InnerText, out index) &&
                        index >= 0 && index < sharedStrings.Count
                            ? sharedStrings[index]
                            : string.Empty;
                }
                else if (type == "b")
                {
                    cell.Kind = SpreadsheetCellKind.Boolean;
                    cell.BooleanValue = value != null && value.InnerText == "1";
                }
                else
                {
                    double number;
                    if (double.TryParse(
                            value == null ? string.Empty : value.InnerText,
                            NumberStyles.Float,
                            CultureInfo.InvariantCulture,
                            out number))
                    {
                        cell.Kind = SpreadsheetCellKind.Number;
                        cell.NumberValue = number;
                    }
                    else
                    {
                        cell.Kind = SpreadsheetCellKind.Text;
                        cell.Text = value == null ? string.Empty : value.InnerText;
                    }
                }
            }
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            List<string> result = new List<string>();
            XmlDocument strings = OpcPackageUtility.ReadXmlPart(archive, "xl/sharedStrings.xml");
            if (strings == null)
                return result;

            List<XmlNode> items = FindAll(strings.DocumentElement, "si");
            for (int i = 0; i < items.Count; i++)
            {
                List<XmlNode> textNodes = FindAll(items[i], "t");
                StringBuilder value = new StringBuilder();
                for (int t = 0; t < textNodes.Count; t++)
                    value.Append(textNodes[t].InnerText);
                result.Add(value.ToString());
            }
            return result;
        }

        private static bool TryParseReference(string reference, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrEmpty(reference))
                return false;

            int index = 0;
            while (index < reference.Length && char.IsLetter(reference[index]))
            {
                column = column * 26 +
                    (char.ToUpperInvariant(reference[index]) - 'A' + 1);
                index++;
            }

            if (column <= 0 || index >= reference.Length)
                return false;

            return int.TryParse(reference.Substring(index), out row) && row > 0;
        }

        private static string GetRelationshipId(XmlNode node)
        {
            if (node == null || node.Attributes == null)
                return string.Empty;

            for (int i = 0; i < node.Attributes.Count; i++)
            {
                XmlAttribute attribute = node.Attributes[i];
                if (attribute.LocalName == "id" &&
                    attribute.NamespaceURI ==
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                {
                    return attribute.Value;
                }
            }
            return string.Empty;
        }

        private static string ReadCoreTitle(ZipArchive archive)
        {
            XmlDocument core = OpcPackageUtility.ReadXmlPart(archive, "docProps/core.xml");
            XmlNode title = core == null ? null : FindFirst(core.DocumentElement, "title");
            return title == null || string.IsNullOrEmpty(title.InnerText)
                ? "Workbook"
                : title.InnerText;
        }

        private static XmlNode FindDirect(XmlNode node, string name)
        {
            if (node == null)
                return null;
            for (int i = 0; i < node.ChildNodes.Count; i++)
                if (node.ChildNodes[i].LocalName == name)
                    return node.ChildNodes[i];
            return null;
        }

        private static XmlNode FindFirst(XmlNode node, string name)
        {
            if (node == null)
                return null;
            if (node.LocalName == name)
                return node;
            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
                XmlNode found = FindFirst(node.ChildNodes[i], name);
                if (found != null)
                    return found;
            }
            return null;
        }

        private static List<XmlNode> FindAll(XmlNode node, string name)
        {
            List<XmlNode> result = new List<XmlNode>();
            AddMatches(node, name, result);
            return result;
        }

        private static void AddMatches(XmlNode node, string name, List<XmlNode> result)
        {
            if (node == null)
                return;
            if (node.LocalName == name)
                result.Add(node);
            for (int i = 0; i < node.ChildNodes.Count; i++)
                AddMatches(node.ChildNodes[i], name, result);
        }

        private static string GetAttribute(XmlNode node, string name)
        {
            if (node == null || node.Attributes == null)
                return string.Empty;
            XmlAttribute direct = node.Attributes[name];
            if (direct != null)
                return direct.Value;
            for (int i = 0; i < node.Attributes.Count; i++)
                if (node.Attributes[i].LocalName == name)
                    return node.Attributes[i].Value;
            return string.Empty;
        }
    }

    internal static class XlsxDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            SpreadsheetDocument document =
                SpreadsheetDocument.CreateNew("XLSX Writer Self Test");
            SpreadsheetSheet first = document.Sheets[0];
            first.SetInput(1, 1, "Name");
            first.SetInput(1, 2, "Value");
            first.SetInput(2, 1, "Alpha");
            first.SetInput(2, 2, "12.5");
            first.SetInput(3, 1, "Enabled");
            first.SetInput(3, 2, "true");
            first.SetInput(4, 1, "Formula");
            first.SetInput(4, 2, "=SUM(B2,7.5)");

            SpreadsheetSheet second = document.AddSheet("한국어 시트");
            second.SetInput(1, 1, "한글 XLSX round-trip");
            second.SetInput(2, 1, "42");

            XlsxWriter.Save(document, outputPath);
            ValidatePackage(outputPath, 2);

            SpreadsheetDocument read = XlsxReader.Read(outputPath);
            if (read.Sheets.Count != 2)
                throw new InvalidOperationException("XLSX sheet count round-trip failed.");

            SpreadsheetCell alpha = read.Sheets[0].GetCell(2, 1, false);
            SpreadsheetCell value = read.Sheets[0].GetCell(2, 2, false);
            SpreadsheetCell formula = read.Sheets[0].GetCell(4, 2, false);

            if (alpha == null || alpha.Text != "Alpha" ||
                value == null || Math.Abs(value.NumberValue - 12.5) > 0.00001 ||
                formula == null || formula.Kind != SpreadsheetCellKind.Formula)
            {
                throw new InvalidOperationException("XLSX cell round-trip failed.");
            }

            read.Sheets[0].SetInput(2, 2, "99.25");
            string secondPath = outputPath + ".roundtrip.xlsx";

            try
            {
                if (File.Exists(secondPath))
                    File.Delete(secondPath);
                XlsxWriter.Save(read, secondPath);
                ValidatePackage(secondPath, 2);
                SpreadsheetDocument verify = XlsxReader.Read(secondPath);
                SpreadsheetCell changed = verify.Sheets[0].GetCell(2, 2, false);
                if (changed == null || Math.Abs(changed.NumberValue - 99.25) > 0.00001)
                    throw new InvalidOperationException("XLSX edit/save/read round-trip failed.");
            }
            finally
            {
                try { if (File.Exists(secondPath)) File.Delete(secondPath); }
                catch { }
            }
        }

        private static void ValidatePackage(string path, int sheets)
        {
            if (!File.Exists(path))
                throw new InvalidOperationException("XLSX output was not created.");

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                Require(archive, "[Content_Types].xml");
                Require(archive, "_rels/.rels");
                Require(archive, "xl/workbook.xml");
                Require(archive, "xl/_rels/workbook.xml.rels");
                Require(archive, "xl/styles.xml");
                Require(archive, "docProps/core.xml");
                Require(archive, "docProps/app.xml");
                for (int i = 1; i <= sheets; i++)
                    Require(archive, "xl/worksheets/sheet" + i.ToString() + ".xml");
            }
        }

        private static void Require(ZipArchive archive, string name)
        {
            if (archive.GetEntry(name) == null)
                throw new InvalidOperationException("Required XLSX part is missing: " + name);
        }
    }

    internal sealed class SpreadsheetEditSession
    {
        public SpreadsheetDocument Document { get; private set; }
        public string FilePath { get; private set; }
        public bool IsDirty { get; private set; }

        private SpreadsheetEditSession(SpreadsheetDocument document)
        {
            if (document == null)
                throw new ArgumentNullException("document");
            Document = document;
            FilePath = string.Empty;
            IsDirty = true;
        }

        public static SpreadsheetEditSession CreateNew(string title)
        {
            return new SpreadsheetEditSession(
                SpreadsheetDocument.CreateNew(title));
        }

        public static SpreadsheetEditSession Open(string path)
        {
            XlsxEditSafetyResult safety = XlsxEditSafety.Analyze(path);
            if (safety == null || !safety.CanEditSafely)
            {
                throw new InvalidOperationException(
                    safety == null || string.IsNullOrEmpty(safety.Warning)
                        ? "This XLSX cannot yet be edited without risking unsupported-content loss."
                        : safety.Warning);
            }

            SpreadsheetEditSession session =
                new SpreadsheetEditSession(XlsxReader.Read(path));
            session.FilePath = Path.GetFullPath(path);
            session.IsDirty = false;
            return session;
        }

        public void MarkDirty()
        {
            IsDirty = true;
        }

        public void Save()
        {
            if (string.IsNullOrEmpty(FilePath))
                throw new InvalidOperationException("Save As is required for a new workbook.");
            XlsxWriter.Save(Document, FilePath);
            IsDirty = false;
        }

        public void SaveAs(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("A destination path is required.", "path");
            if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                path += ".xlsx";
            XlsxWriter.Save(Document, path);
            FilePath = Path.GetFullPath(path);
            IsDirty = false;
        }
    }
}
