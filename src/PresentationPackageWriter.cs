using System;
using System.IO;

namespace PptxViewer
{
    internal static class PresentationPackageWriter
    {
        public static void Save(
            PresentationDocument document,
            string destinationPath)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            if (string.IsNullOrEmpty(destinationPath))
                throw new ArgumentException("Destination path is required.", "destinationPath");

            string destination = Path.GetFullPath(destinationPath);
            string directory = Path.GetDirectoryName(destination);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string stage = destination + ".package-stage";
            string backup = destination + ".package-backup";

            DeleteIfExists(stage);
            DeleteIfExists(backup);

            try
            {
                PptxWriter.Save(document, stage);
                PptxTableWriter.InjectTables(document, stage);
                ReplaceSafely(stage, destination, backup);
            }
            catch
            {
                DeleteIfExists(stage);
                throw;
            }
        }

        private static void ReplaceSafely(
            string stage,
            string destination,
            string backup)
        {
            if (!File.Exists(destination))
            {
                File.Move(stage, destination);
                return;
            }

            File.Move(destination, backup);

            try
            {
                File.Move(stage, destination);
                DeleteIfExists(backup);
            }
            catch
            {
                try
                {
                    DeleteIfExists(destination);
                    if (File.Exists(backup))
                        File.Move(backup, destination);
                }
                catch { }

                throw;
            }
        }

        private static void DeleteIfExists(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch { }
        }
    }
}
