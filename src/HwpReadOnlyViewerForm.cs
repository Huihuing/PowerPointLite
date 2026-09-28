using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class HwpReadOnlyViewerForm : Form
    {
        private readonly HwpReadResult result;
        private readonly RichTextBox documentView;

        public HwpReadOnlyViewerForm(string path)
        {
            result = HwpReader.Read(path);

            Text = Path.GetFileName(path) + " - HWP Read-only";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1050;
            Height = 800;
            MinimumSize = new Size(720, 520);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);

            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 72;
            header.Padding = new Padding(18, 10, 18, 8);
            header.BackColor = ApplicationTheme.Toolbar;
            Controls.Add(header);

            Label title = new Label();
            title.Left = 18;
            title.Top = 10;
            title.Width = 720;
            title.Height = 26;
            title.Text = Path.GetFileName(path);
            title.Font = new Font(Font, FontStyle.Bold);
            title.ForeColor = ApplicationTheme.PrimaryText;
            header.Controls.Add(title);

            Label details = new Label();
            details.Left = 18;
            details.Top = 38;
            details.Width = 900;
            details.Height = 22;
            details.ForeColor = ApplicationTheme.SecondaryText;
            details.Text =
                "HWP 5.x read-only  •  version " + result.Info.Version +
                "  •  " + result.Info.SectionCount.ToString() + " section(s)" +
                (result.Info.IsCompressed ? "  •  compressed" : string.Empty) +
                "  •  no document rewrite";
            header.Controls.Add(details);

            Panel host = new Panel();
            host.Dock = DockStyle.Fill;
            host.Padding = new Padding(42, 28, 42, 28);
            host.BackColor = ApplicationTheme.Canvas;
            Controls.Add(host);
            host.BringToFront();

            documentView = new RichTextBox();
            documentView.Dock = DockStyle.Fill;
            documentView.ReadOnly = true;
            documentView.BorderStyle = BorderStyle.FixedSingle;
            documentView.BackColor = Color.White;
            documentView.ForeColor = Color.FromArgb(32, 36, 42);
            documentView.Font = new Font("Segoe UI", 10.5f);
            documentView.DetectUrls = true;
            host.Controls.Add(documentView);

            Label footer = new Label();
            footer.Dock = DockStyle.Bottom;
            footer.Height = 30;
            footer.Padding = new Padding(12, 0, 12, 0);
            footer.TextAlign = ContentAlignment.MiddleLeft;
            footer.BackColor = ApplicationTheme.Toolbar;
            footer.ForeColor = ApplicationTheme.SecondaryText;
            footer.Text =
                "Experimental public-spec reader: plain paragraph text only. " +
                "Tables, images and full HWP formatting are not claimed yet.";
            Controls.Add(footer);
            footer.BringToFront();

            RenderDocument();
        }

        private void RenderDocument()
        {
            documentView.SuspendLayout();
            documentView.Clear();

            TextDocument document = result.Document;
            if (document != null)
            {
                for (int p = 0; p < document.Paragraphs.Count; p++)
                {
                    DocumentParagraph paragraph = document.Paragraphs[p];
                    if (paragraph != null)
                        documentView.AppendText(paragraph.PlainText ?? string.Empty);
                    if (p + 1 < document.Paragraphs.Count)
                        documentView.AppendText(Environment.NewLine);
                }
            }

            documentView.Select(0, 0);
            documentView.ResumeLayout();
        }
    }
}
