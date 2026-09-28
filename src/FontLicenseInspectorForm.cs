using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class FontLicenseInspectorForm : Form
    {
        private readonly TextBox pathBox;
        private readonly RichTextBox resultBox;

        public FontLicenseInspectorForm()
        {
            Text = "Font License Inspector";
            StartPosition = FormStartPosition.CenterParent;
            Width = 760;
            Height = 620;
            MinimumSize = new Size(620, 480);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);

            Label title = new Label();
            title.Text = "OpenType / TrueType embedding metadata";
            title.Left = 22;
            title.Top = 18;
            title.Width = 680;
            title.Height = 28;
            title.Font = new Font(Font.FontFamily, 14f, FontStyle.Bold);
            title.ForeColor = ApplicationTheme.PrimaryText;
            Controls.Add(title);

            Label warning = new Label();
            warning.Text =
                "fsType metadata is advisory only. The actual font license text controls redistribution, app bundling, and embedding rights.";
            warning.Left = 24;
            warning.Top = 52;
            warning.Width = 690;
            warning.Height = 44;
            warning.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(warning);

            pathBox = new TextBox();
            pathBox.Left = 24;
            pathBox.Top = 106;
            pathBox.Width = 590;
            pathBox.Height = 26;
            pathBox.ReadOnly = true;
            pathBox.BackColor = Color.FromArgb(245, 246, 248);
            pathBox.ForeColor = Color.FromArgb(28, 31, 36);
            Controls.Add(pathBox);

            Button browse = new Button();
            browse.Text = "Browse...";
            browse.Left = 622;
            browse.Top = 104;
            browse.Width = 96;
            browse.Height = 30;
            ApplicationTheme.ApplyButton(browse);
            browse.Click += delegate { BrowseFont(); };
            Controls.Add(browse);

            resultBox = new RichTextBox();
            resultBox.Left = 24;
            resultBox.Top = 148;
            resultBox.Width = 694;
            resultBox.Height = 350;
            resultBox.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Bottom |
                AnchorStyles.Left |
                AnchorStyles.Right;
            resultBox.ReadOnly = true;
            resultBox.BorderStyle = BorderStyle.None;
            resultBox.BackColor = Color.FromArgb(245, 246, 248);
            resultBox.ForeColor = Color.FromArgb(28, 31, 36);
            resultBox.Font = new Font("Consolas", 9.5f);
            Controls.Add(resultBox);

            Label footer = new Label();
            footer.Left = 24;
            footer.Top = 510;
            footer.Width = 580;
            footer.Height = 44;
            footer.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            footer.Text =
                "This tool never adds the selected font to the app or document. " +
                "Embedding remains deny-by-default until the license itself is reviewed.";
            footer.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(footer);

            Button close = new Button();
            close.Text = "Close";
            close.Left = 622;
            close.Top = 516;
            close.Width = 96;
            close.Height = 34;
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            ApplicationTheme.ApplyButton(close);
            close.Click += delegate { Close(); };
            Controls.Add(close);

            AcceptButton = browse;
            CancelButton = close;

            ShowInitialText();
        }

        private void BrowseFont()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter =
                    "OpenType / TrueType fonts (*.ttf;*.otf;*.ttc)|*.ttf;*.otf;*.ttc|All files (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                Inspect(dialog.FileName);
            }
        }

        private void Inspect(string path)
        {
            pathBox.Text = path;

            FontLicenseInfo info = OpenTypeFontLicenseReader.Read(path);
            StringBuilder text = new StringBuilder();

            text.AppendLine("File: " + Path.GetFileName(path));
            text.AppendLine("Metadata available: " + YesNo(info.MetadataAvailable));
            text.AppendLine("OS/2 fsType: 0x" + info.RawFsType.ToString("X4"));
            text.AppendLine("Embedding level: " + info.EmbeddingLevel.ToString());
            text.AppendLine();
            text.AppendLine("Can embed: " + YesNo(info.CanEmbed));
            text.AppendLine("Can embed for editing: " + YesNo(info.CanEmbedForEditing));
            text.AppendLine("Can preview/print: " + YesNo(info.CanPreviewAndPrint));
            text.AppendLine("Can subset: " + YesNo(info.CanSubset));
            text.AppendLine("Bitmap-only embedding: " + YesNo(info.BitmapEmbeddingOnly));
            text.AppendLine("License-text review required: " + YesNo(info.RequiresLicenseTextReview));
            text.AppendLine();
            text.AppendLine("Reader note:");
            text.AppendLine(info.Note ?? string.Empty);
            text.AppendLine();
            text.AppendLine("Project policy:");
            text.AppendLine("- fsType alone does not authorize application bundling or font-file redistribution.");
            text.AppendLine("- Actual license text takes priority over font metadata.");
            text.AppendLine("- Application bundling requires an explicit redistribution/app-embedding license review.");
            text.AppendLine("- Document/PDF font embedding is denied by default until the license is confirmed.");

            resultBox.Text = text.ToString();
        }

        private void ShowInitialText()
        {
            resultBox.Text =
                "Choose a .ttf, .otf or .ttc file to inspect its OS/2 fsType embedding metadata.\r\n\r\n" +
                "Important: a technical 'Installable' or 'Editable' flag is not a substitute for reading the font license. " +
                "This inspector is a development aid, not a legal-license oracle.";
        }

        private static string YesNo(bool value)
        {
            return value ? "Yes" : "No";
        }
    }
}
