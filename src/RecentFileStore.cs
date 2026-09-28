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
internal static class RecentFileStore
    {
        private static string FilePath
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PptxViewer");

                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "recent.txt");
            }
        }

        public static List<string> Get()
        {
            List<string> result = new List<string>();

            try
            {
                if (!File.Exists(FilePath))
                    return result;

                string[] lines = File.ReadAllLines(FilePath, Encoding.UTF8);

                for (int i = 0; i < lines.Length; i++)
                {
                    string path = lines[i].Trim();

                    if (!string.IsNullOrEmpty(path) &&
                        File.Exists(path) &&
                        !result.Contains(path))
                    {
                        result.Add(path);
                    }

                    if (result.Count >= 10)
                        break;
                }
            }
            catch { }

            return result;
        }

        public static void Add(string path)
        {
            if (string.IsNullOrEmpty(path))
                return;

            try
            {
                List<string> items = Get();
                items.RemoveAll(delegate(string p)
                {
                    return string.Equals(
                        p,
                        path,
                        StringComparison.OrdinalIgnoreCase);
                });

                items.Insert(0, path);

                while (items.Count > 10)
                    items.RemoveAt(items.Count - 1);

                File.WriteAllLines(FilePath, items.ToArray(), Encoding.UTF8);
            }
            catch { }
        }

        public static void Clear()
        {
            try
            {
                if (File.Exists(FilePath))
                    File.Delete(FilePath);
            }
            catch { }
        }
    }
}
