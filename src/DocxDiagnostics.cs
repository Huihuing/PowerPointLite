using System;
using System.IO;
using System.IO.Compression;

namespace PptxViewer
{
    internal static class DocxDiagnostics
    {
        public static void CreateAndValidate(string outputPath)
        {
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path is required.", "outputPath");

            TextDocument document = TextDocument.CreateNew("DOCX Writer Self Test");

            DocumentParagraph title = document.AddParagraph(string.Empty);
            title.Alignment = DocumentParagraphAlignment.Center;
            DocumentTextRun titleRun = title.AddRun("DOCX Writer Self Test");
            titleRun.Bold = true;
            titleRun.FontSizePoints = 22f;
            titleRun.ColorHex = "356AE6";

            DocumentParagraph body = document.AddParagraph(string.Empty);
            DocumentTextRun bodyRun = body.AddRun(
                "This DOCX was generated directly from OOXML by the project. ");
            bodyRun.FontSizePoints = 11f;

            DocumentTextRun emphasized = body.AddRun(
                "No Microsoft Word binary, DLL, template, or bundled commercial font is required.");
            emphasized.Bold = true;
            emphasized.Italic = true;
            emphasized.FontSizePoints = 11f;

            DocumentParagraph korean = document.AddParagraph(string.Empty);
            korean.Alignment = DocumentParagraphAlignment.Left;
            DocumentTextRun koreanRun = korean.AddRun(
                "한글 텍스트 및 기본 서식 round-trip 테스트");
            koreanRun.FontSizePoints = 12f;
            koreanRun.Underline = true;

            DocxWriter.Save(document, outputPath);
            ValidatePackage(outputPath);

            TextDocument read = DocxReader.Read(outputPath);
            if (read == null || read.Paragraphs.Count != 3)
                throw new InvalidOperationException("DOCX Reader did not return the expected paragraphs.");

            if (read.Paragraphs[0].PlainText != "DOCX Writer Self Test")
                throw new InvalidOperationException("DOCX title round-trip failed.");

            if (read.Paragraphs[1].Runs.Count != 2 ||
                !read.Paragraphs[1].Runs[1].Bold ||
                !read.Paragraphs[1].Runs[1].Italic)
            {
                throw new InvalidOperationException("DOCX run formatting round-trip failed.");
            }

            if (read.Paragraphs[2].Runs.Count != 1 ||
                !read.Paragraphs[2].Runs[0].Underline)
            {
                throw new InvalidOperationException("DOCX underline round-trip failed.");
            }

            read.Paragraphs[2].Runs[0].Text =
                "한글 텍스트 수정 및 재저장 테스트";

            string secondPath = outputPath + ".roundtrip.docx";

            try
            {
                if (File.Exists(secondPath))
                    File.Delete(secondPath);

                DocxWriter.Save(read, secondPath);
                ValidatePackage(secondPath);

                TextDocument verify = DocxReader.Read(secondPath);
                if (verify.Paragraphs.Count != 3 ||
                    verify.Paragraphs[2].PlainText !=
                        "한글 텍스트 수정 및 재저장 테스트")
                {
                    throw new InvalidOperationException("DOCX edit/save/read round-trip failed.");
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(secondPath))
                        File.Delete(secondPath);
                }
                catch { }
            }
        }

        private static void ValidatePackage(string path)
        {
            if (!File.Exists(path))
                throw new InvalidOperationException("DOCX output was not created.");

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                Require(archive, "[Content_Types].xml");
                Require(archive, "_rels/.rels");
                Require(archive, "word/document.xml");
                Require(archive, "word/styles.xml");
                Require(archive, "word/_rels/document.xml.rels");
                Require(archive, "docProps/core.xml");
                Require(archive, "docProps/app.xml");
            }
        }

        private static void Require(ZipArchive archive, string name)
        {
            if (archive.GetEntry(name) == null)
                throw new InvalidOperationException("Required DOCX part is missing: " + name);
        }
    }
}
