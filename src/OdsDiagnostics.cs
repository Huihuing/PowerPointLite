using System;
using System.IO;
using System.IO.Compression;

namespace PptxViewer
{
    internal static class OdsDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            SpreadsheetDocument document = SpreadsheetDocument.CreateNew("ODS Writer Self Test");
            SpreadsheetSheet first = document.Sheets[0];
            first.SetText(1, 1, "Name");
            first.SetText(1, 2, "Value");
            first.SetText(2, 1, "Alpha");
            first.SetNumber(2, 2, 42.5);
            first.SetBoolean(3, 1, true);
            first.SetFormula(3, 2, "=1+2");

            SpreadsheetSheet second = document.AddSheet("Second");
            second.SetText(1, 1, "ODS second worksheet");

            OdsWriter.Save(document, outputPath);
            ValidatePackage(outputPath);

            OdsEditSafetyResult safety = OdsEditSafety.Analyze(outputPath);
            if (safety == null || !safety.CanEditSafely)
                throw new InvalidOperationException(safety == null ? "ODS safety analysis failed." : safety.Warning);

            SpreadsheetDocument reopened = OdsReader.Read(outputPath);
            if (reopened.Sheets.Count != 2)
                throw new InvalidOperationException("ODS reader worksheet count mismatch.");

            SpreadsheetCell number = reopened.Sheets[0].GetCell(2, 2, false);
            if (number == null || Math.Abs(number.NumberValue - 42.5) > 0.00001)
                throw new InvalidOperationException("ODS numeric cell round-trip failed.");

            reopened.Sheets[0].SetNumber(2, 2, 99.25);
            OdsWriter.Save(reopened, outputPath);

            SpreadsheetDocument verify = OdsReader.Read(outputPath);
            SpreadsheetCell changed = verify.Sheets[0].GetCell(2, 2, false);
            if (changed == null || Math.Abs(changed.NumberValue - 99.25) > 0.00001)
                throw new InvalidOperationException("ODS edit/save/read round-trip failed.");
        }

        private static void ValidatePackage(string path)
        {
            if (!File.Exists(path))
                throw new InvalidOperationException("ODS writer output was not created.");

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                Require(archive, "mimetype");
                Require(archive, "content.xml");
                Require(archive, "styles.xml");
                Require(archive, "meta.xml");
                Require(archive, "META-INF/manifest.xml");

                string mime = OdfPackageUtility.ReadTextEntry(archive, "mimetype").Trim();
                if (!string.Equals(mime, OdfPackageUtility.SpreadsheetMimeType, StringComparison.Ordinal))
                    throw new InvalidOperationException("ODS mimetype is invalid.");
            }
        }

        private static void Require(ZipArchive archive, string path)
        {
            if (archive.GetEntry(path) == null)
                throw new InvalidOperationException("Required ODS part is missing: " + path);
        }
    }
}
