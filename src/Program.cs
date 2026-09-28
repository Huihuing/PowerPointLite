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

                string startupFile = null;
                if (args != null && args.Length > 0 && File.Exists(args[0]))
                    startupFile = args[0];

                using (MainForm form = new MainForm(startupFile))
                    Application.Run(form);
            }
            catch (Exception ex)
            {
                CrashReporter.Report("Startup", ex);
            }
        }

        private static bool TryRunWriterSelfTest(string[] args)
        {
            if (args == null || args.Length == 0 ||
                !string.Equals(args[0], "--writer-selftest", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string outputPath =
                args.Length > 1 && !string.IsNullOrEmpty(args[1])
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "TEST_WRITER_OUTPUT.pptx");

            PptxWriterDiagnostics.CreateAndValidate(outputPath);

            MessageBox.Show(
                "PPTX Writer self-test passed.\r\n\r\n" + outputPath,
                "PowerPointLite Writer Test",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return true;
        }

        private static bool TryRunDocxSelfTest(string[] args)
        {
            if (args == null || args.Length == 0 ||
                !string.Equals(args[0], "--docx-selftest", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string outputPath =
                args.Length > 1 && !string.IsNullOrEmpty(args[1])
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "TEST_DOCX_OUTPUT.docx");

            DocxDiagnostics.CreateAndValidate(outputPath);

            MessageBox.Show(
                "DOCX create/read/edit round-trip self-test passed.\r\n\r\n" + outputPath,
                "PowerPointLite DOCX Test",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return true;
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
