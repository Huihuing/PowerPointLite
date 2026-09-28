using System;
using System.Drawing;
using System.Windows.Forms;

namespace PptxViewer
{
    internal enum WorkspaceAction
    {
        None,
        NewPresentation,
        NewDocx,
        NewXlsx,
        NewHwpx,
        NewOdt,
        NewOds,
        NewOdp,
        OpenExisting,
        ExportPdf,
        ConvertFormat
    }

    internal sealed class OfficeWorkspaceDialog : Form
    {
        public WorkspaceAction SelectedAction { get; private set; }

        public OfficeWorkspaceDialog()
        {
            SelectedAction = WorkspaceAction.None;

            Text = "Document Workspace";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            Width = 820;
            Height = 690;
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);

            Label title = new Label();
            title.Left = 28;
            title.Top = 22;
            title.Width = 740;
            title.Height = 38;
            title.Text = "Create, open, export, or convert a document";
            title.Font = new Font(Font.FontFamily, 17f, FontStyle.Bold);
            title.ForeColor = ApplicationTheme.PrimaryText;
            Controls.Add(title);

            Label subtitle = new Label();
            subtitle.Left = 30;
            subtitle.Top = 62;
            subtitle.Width = 740;
            subtitle.Height = 38;
            subtitle.Text =
                "Independent readers and writers with project-owned UI. " +
                "Microsoft/Hancom application assets are not bundled.";
            subtitle.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(subtitle);

            AddCard(
                "Presentation",
                "PPTX",
                "Slides, text, images, shapes, tables, viewer and presentation mode.",
                30,
                110,
                WorkspaceAction.NewPresentation);

            AddCard(
                "Document",
                "DOCX",
                "Paragraphs and styled text using the shared document model.",
                410,
                110,
                WorkspaceAction.NewDocx);

            AddCard(
                "Spreadsheet",
                "XLSX",
                "Worksheets, cells and formula storage. Calculation engine is still limited.",
                30,
                210,
                WorkspaceAction.NewXlsx);

            AddCard(
                "Korean XML document",
                "HWPX · experimental",
                "Public-format reader/writer. Real Hancom compatibility still requires verification.",
                410,
                210,
                WorkspaceAction.NewHwpx);

            AddCard(
                "OpenDocument text",
                "ODT · experimental",
                "ODF text reader/writer connected to the shared document editor.",
                30,
                310,
                WorkspaceAction.NewOdt);

            AddCard(
                "OpenDocument sheet",
                "ODS · experimental",
                "ODF spreadsheet reader/writer with a lightweight worksheet editor.",
                410,
                310,
                WorkspaceAction.NewOds);

            AddCard(
                "OpenDocument presentation",
                "ODP · experimental",
                "ODF presentation reader/writer with independent slide editing tools.",
                30,
                410,
                WorkspaceAction.NewOdp);

            Panel note = new Panel();
            note.Left = 410;
            note.Top = 410;
            note.Width = 350;
            note.Height = 90;
            note.BackColor = ApplicationTheme.Surface;
            Controls.Add(note);

            Label noteTitle = new Label();
            noteTitle.Left = 16;
            noteTitle.Top = 12;
            noteTitle.Width = 318;
            noteTitle.Height = 22;
            noteTitle.Text = "Safety first";
            noteTitle.Font = new Font(Font, FontStyle.Bold);
            noteTitle.ForeColor = ApplicationTheme.PrimaryText;
            note.Controls.Add(noteTitle);

            Label noteBody = new Label();
            noteBody.Left = 16;
            noteBody.Top = 38;
            noteBody.Width = 318;
            noteBody.Height = 44;
            noteBody.Text =
                "Existing files are edited or converted only when the current parser can represent them without known silent data loss.";
            noteBody.ForeColor = ApplicationTheme.SecondaryText;
            note.Controls.Add(noteBody);

            Button open = new Button();
            open.Text = "Open existing editable file...";
            open.Left = 30;
            open.Top = 528;
            open.Width = 350;
            open.Height = 38;
            ApplicationTheme.ApplyButton(open);
            open.Click += delegate
            {
                SelectedAction = WorkspaceAction.OpenExisting;
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(open);

            Button exportPdf = new Button();
            exportPdf.Text = "Export to PDF...";
            exportPdf.Left = 410;
            exportPdf.Top = 528;
            exportPdf.Width = 165;
            exportPdf.Height = 38;
            ApplicationTheme.ApplyButton(exportPdf);
            exportPdf.Click += delegate
            {
                SelectedAction = WorkspaceAction.ExportPdf;
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(exportPdf);

            Button convert = new Button();
            convert.Text = "Convert format...";
            convert.Left = 585;
            convert.Top = 528;
            convert.Width = 175;
            convert.Height = 38;
            ApplicationTheme.ApplyButton(convert);
            convert.Click += delegate
            {
                SelectedAction = WorkspaceAction.ConvertFormat;
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(convert);

            Button cancel = new Button();
            cancel.Text = "Cancel";
            cancel.Left = 660;
            cancel.Top = 576;
            cancel.Width = 100;
            cancel.Height = 38;
            ApplicationTheme.ApplyButton(cancel);
            cancel.Click += delegate
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };
            Controls.Add(cancel);

            CancelButton = cancel;
        }

        private void AddCard(
            string titleText,
            string formatText,
            string description,
            int left,
            int top,
            WorkspaceAction action)
        {
            Panel card = new Panel();
            card.Left = left;
            card.Top = top;
            card.Width = 350;
            card.Height = 90;
            card.BackColor = ApplicationTheme.Surface;
            card.Cursor = Cursors.Hand;
            card.TabStop = true;

            Label title = new Label();
            title.Left = 16;
            title.Top = 11;
            title.Width = 210;
            title.Height = 23;
            title.Text = titleText;
            title.Font = new Font(Font, FontStyle.Bold);
            title.ForeColor = ApplicationTheme.PrimaryText;
            title.Cursor = Cursors.Hand;
            card.Controls.Add(title);

            Label format = new Label();
            format.Left = 224;
            format.Top = 11;
            format.Width = 110;
            format.Height = 23;
            format.TextAlign = ContentAlignment.MiddleRight;
            format.Text = formatText;
            format.ForeColor = ApplicationTheme.Accent;
            format.Cursor = Cursors.Hand;
            card.Controls.Add(format);

            Label body = new Label();
            body.Left = 16;
            body.Top = 39;
            body.Width = 318;
            body.Height = 43;
            body.Text = description;
            body.ForeColor = ApplicationTheme.SecondaryText;
            body.Cursor = Cursors.Hand;
            card.Controls.Add(body);

            EventHandler choose = delegate
            {
                SelectedAction = action;
                DialogResult = DialogResult.OK;
                Close();
            };

            card.Click += choose;
            title.Click += choose;
            format.Click += choose;
            body.Click += choose;

            card.MouseEnter += delegate { card.BackColor = ApplicationTheme.SurfaceHover; };
            card.MouseLeave += delegate { card.BackColor = ApplicationTheme.Surface; };

            Controls.Add(card);
        }
    }
}
