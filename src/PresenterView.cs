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
internal sealed class PresenterViewForm : Form
    {
        private readonly PictureBox currentPicture;
        private readonly PictureBox nextPicture;
        private readonly RichTextBox notes;
        private readonly Label slideLabel;
        private readonly Label titleLabel;
        private readonly Label timerLabel;
        private readonly System.Windows.Forms.Timer timer;
        private DateTime startedAt;

        public event EventHandler PreviousRequested;
        public event EventHandler NextRequested;
        public event Action<int> GoToSlideRequested;
        public event EventHandler StartShowRequested;

        public PresenterViewForm()
        {
            Text = "Presenter View";
            StartPosition = FormStartPosition.Manual;
            Width = 1200;
            Height = 760;
            MinimumSize = new Size(900, 600);
            BackColor = ApplicationTheme.Window;
            ForeColor = ApplicationTheme.PrimaryText;

            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 2;
            root.RowCount = 3;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68f));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56f));
            root.BackColor = ApplicationTheme.Divider;
            root.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 62f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 38f));
            Controls.Add(root);

            Panel top = new Panel();
            top.Dock = DockStyle.Fill;
            top.BackColor = ApplicationTheme.Toolbar;
            root.Controls.Add(top, 0, 0);
            root.SetColumnSpan(top, 2);

            Button prev = MakeButton("Previous", 8);
            prev.Click += delegate
            {
                if (PreviousRequested != null)
                    PreviousRequested(this, EventArgs.Empty);
            };
            top.Controls.Add(prev);

            Button next = MakeButton("Next", 98);
            next.Click += delegate
            {
                if (NextRequested != null)
                    NextRequested(this, EventArgs.Empty);
            };
            top.Controls.Add(next);

            Button go = MakeButton("Go To", 188);
            go.Click += delegate
            {
                string value = ViewerDialogs.Prompt(
                    this,
                    "Go to slide",
                    "Slide number:",
                    "");

                int n;

                if (int.TryParse(value, out n) &&
                    GoToSlideRequested != null)
                {
                    GoToSlideRequested(n);
                }
            };
            top.Controls.Add(go);

            Button start = MakeButton("Start Show", 284);
            start.Tag = "F5 Show";
            start.Width = 104;
            ApplicationTheme.ApplyButton(start);
            start.Click += delegate
            {
                if (StartShowRequested != null)
                    StartShowRequested(this, EventArgs.Empty);
            };
            top.Controls.Add(start);

            timerLabel = new Label();
            timerLabel.AutoSize = false;
            timerLabel.Width = 190;
            timerLabel.Height = 32;
            timerLabel.ForeColor = Color.FromArgb(126, 214, 160);
            timerLabel.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
            timerLabel.TextAlign = ContentAlignment.MiddleRight;
            timerLabel.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Right;
            timerLabel.Top = 12;
            top.Controls.Add(timerLabel);

            slideLabel = new Label();
            slideLabel.AutoSize = false;
            slideLabel.Width = 104;
            slideLabel.Height = 32;
            slideLabel.ForeColor = ApplicationTheme.PrimaryText;
            slideLabel.Font = new Font("Segoe UI", 11f, FontStyle.Bold);
            slideLabel.TextAlign = ContentAlignment.MiddleRight;
            slideLabel.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Right;
            slideLabel.Top = 12;
            top.Controls.Add(slideLabel);

            Action layoutTopBar =
                delegate
                {
                    int gap = 8;
                    int x = 12;

                    prev.Left = x;
                    x +=
                        prev.Width +
                        gap;

                    next.Left = x;
                    x +=
                        next.Width +
                        gap;

                    go.Left = x;
                    x +=
                        go.Width +
                        gap;

                    start.Left = x;

                    slideLabel.Left =
                        Math.Max(
                            start.Right + 20,
                            top.ClientSize.Width -
                            slideLabel.Width -
                            14);

                    timerLabel.Left =
                        Math.Max(
                            start.Right + 12,
                            slideLabel.Left -
                            timerLabel.Width -
                            12);

                    bool enoughStatusRoom =
                        timerLabel.Left >
                        start.Right + 8;

                    timerLabel.Visible =
                        enoughStatusRoom;
                };

            top.Resize += delegate
            {
                layoutTopBar();
            };

            layoutTopBar();

            currentPicture = new PictureBox();
            currentPicture.Dock = DockStyle.Fill;
            currentPicture.SizeMode = PictureBoxSizeMode.Zoom;
            currentPicture.BackColor = ApplicationTheme.Canvas;
            currentPicture.BorderStyle = BorderStyle.FixedSingle;
            root.Controls.Add(currentPicture, 0, 1);

            Panel nextPanel = new Panel();
            nextPanel.Dock = DockStyle.Fill;
            nextPanel.Padding = new Padding(12);
            nextPanel.BackColor = ApplicationTheme.Surface;
            root.Controls.Add(nextPanel, 1, 1);

            Label nextLabel = new Label();
            nextLabel.Dock = DockStyle.Top;
            nextLabel.Height = 30;
            nextLabel.Text = "Next slide";
            nextLabel.ForeColor = ApplicationTheme.SecondaryText;
            nextLabel.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            nextPanel.Controls.Add(nextLabel);

            nextPicture = new PictureBox();
            nextPicture.Dock = DockStyle.Fill;
            nextPicture.SizeMode = PictureBoxSizeMode.Zoom;
            nextPicture.BackColor = ApplicationTheme.Canvas;
            nextPicture.BorderStyle = BorderStyle.FixedSingle;
            nextPanel.Controls.Add(nextPicture);
            nextPicture.BringToFront();

            Panel notesPanel = new Panel();
            notesPanel.Dock = DockStyle.Fill;
            notesPanel.Padding = new Padding(12);
            notesPanel.BackColor = ApplicationTheme.Window;
            root.Controls.Add(notesPanel, 0, 2);
            root.SetColumnSpan(notesPanel, 2);

            titleLabel = new Label();
            titleLabel.Dock = DockStyle.Top;
            titleLabel.Height = 34;
            titleLabel.ForeColor = ApplicationTheme.PrimaryText;
            titleLabel.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
            notesPanel.Controls.Add(titleLabel);

            notes = new RichTextBox();
            notes.Dock = DockStyle.Fill;
            notes.ReadOnly = true;
            notes.BackColor = ApplicationTheme.Surface;
            notes.ForeColor = ApplicationTheme.PrimaryText;
            notes.Font = new Font("Segoe UI", 12f);
            notes.BorderStyle = BorderStyle.FixedSingle;
            notesPanel.Controls.Add(notes);
            notes.BringToFront();

            startedAt = DateTime.Now;

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000;
            timer.Tick += delegate
            {
                TimeSpan elapsed = DateTime.Now - startedAt;
                timerLabel.Text = "Elapsed " + elapsed.ToString(@"hh\:mm\:ss");
            };
            timer.Start();

            FormClosed += delegate
            {
                timer.Stop();

                if (currentPicture.Image != null)
                    currentPicture.Image.Dispose();

                if (nextPicture.Image != null)
                    nextPicture.Image.Dispose();
            };
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            PositionOnPreferredScreen();
        }

        private void PositionOnPreferredScreen()
        {
            Screen ownerScreen =
                Owner != null
                ? Screen.FromControl(Owner)
                : Screen.FromPoint(Cursor.Position);

            Screen targetScreen =
                ownerScreen;

            Screen[] screens =
                Screen.AllScreens;

            if (screens != null &&
                screens.Length > 1)
            {
                for (int i = 0;
                     i < screens.Length;
                     i++)
                {
                    if (!string.Equals(
                            screens[i].DeviceName,
                            ownerScreen.DeviceName,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        targetScreen =
                            screens[i];
                        break;
                    }
                }
            }

            Rectangle area =
                targetScreen.WorkingArea;

            int width =
                Math.Max(
                    400,
                    Math.Min(
                        Width,
                        area.Width));

            int height =
                Math.Max(
                    300,
                    Math.Min(
                        Height,
                        area.Height));

            StartPosition =
                FormStartPosition.Manual;

            Bounds =
                new Rectangle(
                    area.Left +
                        Math.Max(
                            0,
                            (area.Width - width) / 2),
                    area.Top +
                        Math.Max(
                            0,
                            (area.Height - height) / 2),
                    width,
                    height);
        }

        private static Button MakeButton(string text, int left)
        {
            Button b = new Button();
            b.Text = text;
            b.Left = left;
            b.Top = 12;
            b.Width = 84;
            b.Height = 32;
            b.Tag = text;
            ApplicationTheme.ApplyButton(b);
            return b;
        }

        public void UpdateSlide(
            Image current,
            Image next,
            int slideNumber,
            int totalSlides,
            string title,
            string speakerNotes)
        {
            Image oldCurrent = currentPicture.Image;
            currentPicture.Image =
                current == null
                ? null
                : new Bitmap(current);

            if (oldCurrent != null)
                oldCurrent.Dispose();

            Image oldNext = nextPicture.Image;
            nextPicture.Image =
                next == null
                ? null
                : new Bitmap(next);

            if (oldNext != null)
                oldNext.Dispose();

            slideLabel.Text =
                slideNumber.ToString() +
                " / " +
                totalSlides.ToString();

            titleLabel.Text =
                string.IsNullOrWhiteSpace(title)
                ? "Speaker Notes"
                : title;

            notes.Text = speakerNotes ?? "";
        }
    }
}
