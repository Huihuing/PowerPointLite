using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace PptxViewer
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate(object sender, System.Threading.ThreadExceptionEventArgs e)
            {
                CrashReporter.Report("UI thread", e.Exception);
            };

            AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
            {
                Exception ex = e.ExceptionObject as Exception;
                CrashReporter.Report("Unhandled AppDomain", ex ?? new Exception("Unknown fatal error"));
            };

            try
            {
                CrashReporter.WriteLine("Starting PowerPointLite 1.3");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                if (TryRunWriterSelfTest(args))
                    return;
                if (TryRunDocxSelfTest(args))
                    return;
                if (TryRunXlsxSelfTest(args))
                    return;
                if (TryRunHwpxSelfTest(args))
                    return;
                if (TryRunHwpParserSelfTest(args))
                    return;
                if (TryRunOdtSelfTest(args))
                    return;
                if (TryRunOdsSelfTest(args))
                    return;
                if (TryRunOdpSelfTest(args))
                    return;
                if (TryRunPdfSelfTest(args))
                    return;
                if (TryRunConversionSelfTest(args))
                    return;
                if (TryRunFontLicenseSelfTest(args))
                    return;

                string startupFile = null;
                if (args != null && args.Length > 0 && File.Exists(args[0]))
                    startupFile = args[0];

                using (MainForm form = new MainForm(startupFile))
                    Application.Run(form);
            }
            catch (Exception ex)
            {
                if (IsSelfTestCommand(args))
                {
                    Environment.ExitCode = 1;
                    CrashReporter.WriteLine("SELF-TEST FAILED: " + ex.ToString());
                    return;
                }

                Environment.ExitCode = 1;
                CrashReporter.Report("Startup", ex);
            }
        }

        private static bool IsSelfTestCommand(string[] args)
        {
            if (args == null || args.Length == 0 || string.IsNullOrEmpty(args[0]))
                return false;

            return string.Equals(args[0], "--writer-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--docx-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--xlsx-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--hwpx-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--hwp-parser-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--odt-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--ods-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--odp-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--pdf-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--conversion-selftest", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(args[0], "--font-license-selftest", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryRunWriterSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--writer-selftest"))
                return false;

            string outputPath = ResolveOutputFile(args, "TEST_WRITER_OUTPUT.pptx");
            PptxWriterDiagnostics.CreateAndValidate(outputPath);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("PPTX Writer self-test passed: " + outputPath);
            return true;
        }

        private static bool TryRunDocxSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--docx-selftest"))
                return false;

            string outputPath = ResolveOutputFile(args, "TEST_DOCX_OUTPUT.docx");
            DocxDiagnostics.CreateAndValidate(outputPath);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("DOCX self-test passed: " + outputPath);
            return true;
        }

        private static bool TryRunXlsxSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--xlsx-selftest"))
                return false;

            string outputPath = ResolveOutputFile(args, "TEST_XLSX_OUTPUT.xlsx");
            XlsxDiagnostics.CreateAndValidate(outputPath);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("XLSX self-test passed: " + outputPath);
            return true;
        }

        private static bool TryRunHwpxSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--hwpx-selftest"))
                return false;

            string outputPath = ResolveOutputFile(args, "TEST_HWPX_OUTPUT.hwpx");
            HwpxDiagnostics.CreateAndValidate(outputPath);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("HWPX structural self-test passed: " + outputPath);
            return true;
        }

        private static bool TryRunHwpParserSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--hwp-parser-selftest"))
                return false;

            HwpDiagnostics.ValidateParserFoundation();
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("HWP FileHeader/record parser self-test passed.");
            return true;
        }

        private static bool TryRunOdtSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--odt-selftest"))
                return false;

            string outputPath = ResolveOutputFile(args, "TEST_ODT_OUTPUT.odt");
            OdtDiagnostics.CreateAndValidate(outputPath);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("ODT self-test passed: " + outputPath);
            return true;
        }

        private static bool TryRunOdsSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--ods-selftest"))
                return false;

            string outputPath = ResolveOutputFile(args, "TEST_ODS_OUTPUT.ods");
            OdsDiagnostics.CreateAndValidate(outputPath);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("ODS self-test passed: " + outputPath);
            return true;
        }

        private static bool TryRunOdpSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--odp-selftest"))
                return false;

            string outputPath = ResolveOutputFile(args, "TEST_ODP_OUTPUT.odp");
            OdpDiagnostics.CreateAndValidate(outputPath);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("ODP self-test passed: " + outputPath);
            return true;
        }

        private static bool TryRunPdfSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--pdf-selftest"))
                return false;

            string outputDirectory =
                args != null && args.Length > 1 && !string.IsNullOrEmpty(args[1])
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "PDF_SELFTEST_OUTPUT");

            PdfExportDiagnostics.CreateAndValidate(outputDirectory);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("PDF raster export self-test passed: " + outputDirectory);
            return true;
        }

        private static bool TryRunConversionSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--conversion-selftest"))
                return false;

            string outputDirectory =
                args != null && args.Length > 1 && !string.IsNullOrEmpty(args[1])
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "CONVERSION_SELFTEST_OUTPUT");

            DocumentConversionDiagnostics.CreateAndValidate(outputDirectory);
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("Cross-format conversion self-test passed: " + outputDirectory);
            return true;
        }

        private static bool TryRunFontLicenseSelfTest(string[] args)
        {
            if (!MatchesCommand(args, "--font-license-selftest"))
                return false;

            FontLicenseDiagnostics.ValidatePolicyAndParser();
            Environment.ExitCode = 0;
            CrashReporter.WriteLine("Font licensing parser/policy self-test passed.");
            return true;
        }

        private static bool MatchesCommand(string[] args, string command)
        {
            return args != null &&
                args.Length > 0 &&
                string.Equals(
                    args[0],
                    command,
                    StringComparison.OrdinalIgnoreCase);
        }

        private static string ResolveOutputFile(
            string[] args,
            string defaultFileName)
        {
            return args != null &&
                args.Length > 1 &&
                !string.IsNullOrEmpty(args[1])
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        defaultFileName);
        }
    }

    internal static class CrashReporter
    {
        private static string LogDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PptxViewer",
                    "logs");
            }
        }

        public static string LogPath
        {
            get { return Path.Combine(LogDirectory, "last-startup.log"); }
        }

        public static void WriteLine(string text)
        {
            try
            {
                Directory.CreateDirectory(LogDirectory);
                File.AppendAllText(
                    LogPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + text + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch { }
        }

        public static void Report(string stage, Exception ex)
        {
            try
            {
                Directory.CreateDirectory(LogDirectory);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("PowerPointLite fatal error");
                sb.AppendLine("Stage: " + stage);
                sb.AppendLine("Time: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                sb.AppendLine();
                sb.AppendLine(ex == null ? "Unknown error" : ex.ToString());

                File.WriteAllText(LogPath, sb.ToString(), Encoding.UTF8);
            }
            catch { }

            try
            {
                MessageBox.Show(
                    "PowerPointLite could not start correctly.\r\n\r\n" +
                    (ex == null ? "Unknown error" : ex.Message) +
                    "\r\n\r\nDiagnostic log:\r\n" + LogPath,
                    "PowerPointLite startup error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch { }
        }
    }
}
