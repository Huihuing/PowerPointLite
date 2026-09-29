using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace PptxViewer
{
    internal static class DocumentSafetyDiagnostics
    {
        public static void ValidateDenyByDefault(string outputDirectory)
        {
            if (string.IsNullOrEmpty(outputDirectory))
                throw new ArgumentException("Output directory is required.", "outputDirectory");

            string root = Path.GetFullPath(outputDirectory);

            if (Directory.Exists(root))
                Directory.Delete(root, true);

            Directory.CreateDirectory(root);

            ValidateDocx(Path.Combine(root, "UNSAFE_DOCX.docx"));
            ValidateHwpx(Path.Combine(root, "UNSAFE_HWPX.hwpx"));
            ValidateOdt(Path.Combine(root, "UNSAFE_ODT.odt"));
            ValidateOds(Path.Combine(root, "UNSAFE_ODS.ods"));
            ValidateOdp(Path.Combine(root, "UNSAFE_ODP.odp"));
        }

        private static void ValidateDocx(string path)
        {
            DocxDiagnostics.CreateAndValidate(path);
            AddUnknownPart(path);

            DocxEditSafetyResult safety = DocxEditSafety.Analyze(path);
            RequireDenied(
                safety != null && safety.CanEditSafely,
                "DOCX");
            RequireConversionRejected(path, "DOCX");
        }

        private static void ValidateHwpx(string path)
        {
            HwpxDiagnostics.CreateAndValidate(path);
            AddUnknownPart(path);

            HwpxEditSafetyResult safety = HwpxEditSafety.Analyze(path);
            RequireDenied(
                safety != null && safety.CanEditSafely,
                "HWPX");
            RequireConversionRejected(path, "HWPX");
        }

        private static void ValidateOdt(string path)
        {
            OdtDiagnostics.CreateAndValidate(path);
            AddUnknownPart(path);

            OdtEditSafetyResult safety = OdtEditSafety.Analyze(path);
            RequireDenied(
                safety != null && safety.CanEditSafely,
                "ODT");
            RequireConversionRejected(path, "ODT");
        }

        private static void ValidateOds(string path)
        {
            OdsDiagnostics.CreateAndValidate(path);
            AddUnknownPart(path);

            OdsEditSafetyResult safety = OdsEditSafety.Analyze(path);
            RequireDenied(
                safety != null && safety.CanEditSafely,
                "ODS");
            RequireConversionRejected(path, "ODS");
        }

        private static void ValidateOdp(string path)
        {
            OdpDiagnostics.CreateAndValidate(path);
            AddUnknownPart(path);

            OdpEditSafetyResult safety = OdpEditSafety.Analyze(path);
            RequireDenied(
                safety != null && safety.CanEditSafely,
                "ODP");
            RequireConversionRejected(path, "ODP");
        }

        private static void AddUnknownPart(string path)
        {
            using (FileStream stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None))
            using (ZipArchive archive = new ZipArchive(
                stream,
                ZipArchiveMode.Update,
                false))
            {
                ZipArchiveEntry existing =
                    archive.GetEntry("Foreign/unsupported-test.bin");

                if (existing != null)
                    existing.Delete();

                ZipArchiveEntry entry = archive.CreateEntry(
                    "Foreign/unsupported-test.bin",
                    CompressionLevel.Optimal);

                using (Stream output = entry.Open())
                {
                    byte[] bytes = Encoding.ASCII.GetBytes(
                        "PowerPointLite deny-by-default safety fixture");
                    output.Write(bytes, 0, bytes.Length);
                }
            }
        }

        private static void RequireDenied(bool canEditSafely, string format)
        {
            if (canEditSafely)
            {
                throw new InvalidOperationException(
                    format +
                    " safety guard accepted a package containing an unknown foreign part.");
            }
        }

        private static void RequireConversionRejected(string path, string format)
        {
            bool rejected = false;

            try
            {
                SafeDocumentConversion.Load(path);
            }
            catch (InvalidOperationException)
            {
                rejected = true;
            }

            if (!rejected)
            {
                throw new InvalidOperationException(
                    format +
                    " conversion accepted an unsafe package. Conversion must stay deny-by-default when unknown content would be lost.");
            }
        }
    }
}
