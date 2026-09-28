using System;
using System.Drawing;
using System.IO;
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
        OpenExisting
    }

    internal sealed class OfficeWorkspaceDialog : Form
    {
        public WorkspaceAction SelectedAction { get; private set; }

        public OfficeWorkspaceDialog()
        {
            SelectedAction = WorkspaceAction.None;

            Text = "New / Open";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            Width = 720;
            Height = 470;
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;
            Font = new Font("Segoe UI", 9f);

            Label title = new Label();
            title.Left = 28;
            title.Top = 24;
            title.Width = 640;
            title.Height = 36;
            title.Text = "Create or open a document";
            title.Font = new Font(Font.FontFamily, 17f, FontStyle.Bold);
            title.ForeColor = ApplicationTheme.PrimaryText;
            Controls.Add(title);

            Label subtitle = new Label();
            subtitle.Left = 30;
            subtitle.Top = 62;
            subtitle.Width = 640;
            subtitle.Height = 34;
            subtitle.Text = "Independent document tools. No Microsoft or Hancom application assets are bundled.";
            subtitle.ForeColor = ApplicationTheme.SecondaryText;
            Controls.Add(subtitle);

            AddCard(
                "Presentation",
                "PPTX",
                "Slides, text, images, shapes, tables and presentation mode.",
                30,
                112,
                WorkspaceAction.NewPresentation);

            AddCard(
                "Document",
                "DOCX",
                "Paragraphs and styled text using the shared document model.",
                360,
                112,
                WorkspaceAction.NewDocx);

            AddCard(
                "Spreadsheet",
                "XLSX",
                "Worksheets, text, numbers, booleans and formulas.",
                30,
                232,
                WorkspaceAction.NewXlsx);

            AddCard(
                "Open document format",
                "HWPX · experimental",
                "Public-format clean-room reader/writer. Real Hancom compatibility still requires verification.",
                360,
                232,
                WorkspaceAction.NewHwpx);

            Button open = new Button();
            open.Text = "Open existing editable file...";
            open.Left = 30;
            open.Top = 366;
            open.Width = 330;
            open.Height = 38;
            ApplicationTheme.ApplyButton(open);
            open.Click += delegate
            {
                SelectedAction = WorkspaceAction.OpenExisting;
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(open);

            Button cancel = new Button();
            cancel.Text = "Cancel";
            cancel.Left = 570;
            cancel.Top = 366;
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
            card.Width = 310;
            card.Height = 100;
            card.BackColor = ApplicationTheme.Surface;
            card.Cursor = Cursors.Hand;
            card.TabStop = true;

            Label title = new Label();
            title.Left = 16;
            title.Top = 12;
            title.Width = 180;
            title.Height = 24;
            title.Text = titleText;
            title.Font = new Font(Font, FontStyle.Bold);
            title.ForeColor = ApplicationTheme.PrimaryText;
            title.Cursor = Cursors.Hand;
            card.Controls.Add(title);

            Label format = new Label();
            format.Left = 202;
            format.Top = 12;
            format.Width = 92;
            format.Height = 24;
            format.TextAlign = ContentAlignment.MiddleRight;
            format.Text = formatText;
            format.ForeColor = ApplicationTheme.Accent;
            format.Cursor = Cursors.Hand;
            card.Controls.Add(format);

            Label body = new Label();
            body.Left = 16;
            body.Top = 42;
            body.Width = 278;
            body.Height = 46;
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
