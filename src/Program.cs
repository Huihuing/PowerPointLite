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
