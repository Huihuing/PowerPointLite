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
public sealed partial class MainForm : Form
    {
        private void Next()
        {
            if (internalSlideShowMode && TryAdvanceInternalAnimation(true))
                return;

            int target = internalSlideShowMode
                ? FindVisibleSlide(currentIndex + 1, 1)
                : currentIndex + 1;

            if (target >= 0 && target < renderedSlides.Count)
            {
                if (internalSlideShowMode && fullscreen)
                    ShowSlideWithTransition(target, true);
                else
                    ShowSlide(target);
            }
            else if (internalSlideShowMode && fullscreen)
            {
                if (InternalSlideShowLoops())
                {
                    int first =
                        FindVisibleSlide(0, 1);

                    if (first >= 0)
                    {
                        ShowSlideWithTransition(
                            first,
                            true);
                    }
                }
                else
                {
                    EndInternalSlideshow();
                }
            }
        }

        private void ToggleFullscreen()
        {
            if (!fullscreen)
            {
                oldState = WindowState;
                oldBorder = FormBorderStyle;
                sidebarWasCollapsedBeforeFullscreen = mainSplit.Panel1Collapsed;

                mainSplit.Panel1Collapsed = true;
                toolbar.Visible = false;
                status.Visible = false;

                FormBorderStyle = FormBorderStyle.None;
                WindowState = FormWindowState.Maximized;
                fullscreen = true;
            }
            else
            {
                FormBorderStyle = oldBorder;
                WindowState = oldState;

                toolbar.Visible = true;
                status.Visible = true;
                mainSplit.Panel1Collapsed = sidebarWasCollapsedBeforeFullscreen;

                fullscreen = false;
                internalSlideShowMode = false;
                animationRevealCount = 0;
            }

            if (fitMode)
                BeginInvoke((MethodInvoker)ApplyFit);
        }

        private int FindVisibleSlide(int start, int direction)
        {
            if (renderedSlides.Count == 0)
                return -1;

            int i = start;

            while (i >= 0 && i < renderedSlides.Count)
            {
                bool hidden = i < hiddenSlides.Count && hiddenSlides[i];

                if (!hidden)
                    return i;

                i += direction;
            }

            return -1;
        }

        private void ShowNextHiddenSlide()
        {
            if (!internalSlideShowMode)
                return;

            for (int i = currentIndex + 1; i < renderedSlides.Count; i++)
            {
                if (i < hiddenSlides.Count && hiddenSlides[i])
                {
                    ShowSlideWithTransition(i, true);
                    return;
                }

                if (i < hiddenSlides.Count && !hiddenSlides[i])
                    break;
            }
        }

        private void UpdateNotesBox()
        {
            if (currentIndex >= 0 && currentIndex < speakerNotes.Count)
                notesBox.Text = speakerNotes[currentIndex] ?? "";
            else
                notesBox.Clear();
        }

        private void ToggleNotes()
        {
            notesBox.Visible = !notesBox.Visible;
            UpdateNotesBox();

            if (fitMode)
                BeginInvoke((MethodInvoker)ApplyFit);
        }

        private void ToggleBlankScreen(Color color)
        {
            if (!internalSlideShowMode || viewer.Image == null)
                return;

            if (blankScreenActive && blankScreenColor.ToArgb() == color.ToArgb())
            {
                RestoreBlankScreen();
                return;
            }

            RestoreBlankScreen();

            blankScreenBackup = new Bitmap(viewer.Image);
            blankScreenColor = color;
            blankScreenActive = true;

            Bitmap blank = new Bitmap(
                Math.Max(1, viewer.Image.Width),
                Math.Max(1, viewer.Image.Height));

            using (Graphics g = Graphics.FromImage(blank))
                g.Clear(color);

            ReplaceViewerImage(blank);

            if (fitMode) ApplyFit();
            else ApplyZoom();
        }

        private void RestoreBlankScreen()
        {
            if (!blankScreenActive)
                return;

            blankScreenActive = false;

            if (blankScreenBackup != null)
            {
                Image backup = blankScreenBackup;
                blankScreenBackup = null;
                ReplaceViewerImage(backup);
            }
        }

        private void EndInternalSlideshow()
        {
            RestoreBlankScreen();

            if (fullscreen)
                ToggleFullscreen();

            internalSlideShowMode = false;
            animationRevealCount = 0;
            slideNumberBuffer = "";
        }

        private void ShowRecentFilesMenu(Control anchor)
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            List<string> recent = RecentFileStore.Get();

            if (recent.Count == 0)
            {
                ToolStripMenuItem none = new ToolStripMenuItem("(No recent files)");
                none.Enabled = false;
                menu.Items.Add(none);
            }
            else
            {
                for (int i = 0; i < recent.Count; i++)
                {
                    string path = recent[i];
                    ToolStripMenuItem item = new ToolStripMenuItem(
                        Path.GetFileName(path));

                    item.ToolTipText = path;
                    item.Click += delegate
                    {
                        if (File.Exists(path))
                            LoadPresentation(path);
                    };

                    menu.Items.Add(item);
                }

                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add("Clear recent files", null, delegate
                {
                    RecentFileStore.Clear();
                });
            }

            ApplicationTheme.ApplyContextMenu(menu);
            menu.Show(anchor, new Point(0, anchor.Height));
        }

        private void ShowViewMenu(Control anchor)
        {
            ContextMenuStrip menu = new ContextMenuStrip();

            ToolStripMenuItem notes = new ToolStripMenuItem("Speaker Notes");
            notes.Checked = notesBox.Visible;
            notes.Click += delegate { ToggleNotes(); };
            menu.Items.Add(notes);

            ToolStripMenuItem presenter = new ToolStripMenuItem("Presenter View");
            presenter.Checked = presenterView != null && !presenterView.IsDisposed;
            presenter.Click += delegate { TogglePresenterView(); };
            menu.Items.Add(presenter);

            ToolStripMenuItem autoToc = new ToolStripMenuItem("Auto-hide slide thumbnails");
            autoToc.Checked = autoHideSidebar;
            autoToc.Click += delegate
            {
                autoHideSidebar = !autoHideSidebar;

                if (autoHideSidebar && !fullscreen)
                    mainSplit.Panel1Collapsed = true;
                else if (!fullscreen)
                    mainSplit.Panel1Collapsed = false;

                if (fitMode)
                    BeginInvoke((MethodInvoker)ApplyFit);
            };
            menu.Items.Add(autoToc);

            menu.Items.Add("Slide Sorter", null, delegate { ShowSlideSorter(); });
            menu.Items.Add("Find...", null, delegate { FindSlidesDialog(); });
            menu.Items.Add("Go to slide...", null, delegate { GoToSlideDialog(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Fit to window", null, delegate
            {
                fitMode = true;
                ApplyFit();
            });
            menu.Items.Add("100%", null, delegate
            {
                fitMode = false;
                SetZoom(1.0f);
            });
            menu.Items.Add("Full screen", null, delegate { ToggleFullscreen(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Presentation info", null, delegate { ShowPresentationInfo(); });

            ApplicationTheme.ApplyContextMenu(menu);
            menu.Show(anchor, new Point(0, anchor.Height));
        }

        private void GoToSlideDialog()
        {
            string value = ViewerDialogs.Prompt(
                this,
                "Go to slide",
                "Slide number (1-" + renderedSlides.Count.ToString() + "):",
                currentIndex >= 0 ? (currentIndex + 1).ToString() : "1");

            int slideNumber;

            if (int.TryParse(value, out slideNumber) &&
                slideNumber >= 1 &&
                slideNumber <= renderedSlides.Count)
            {
                if (internalSlideShowMode && fullscreen)
                    ShowSlideWithTransition(slideNumber - 1, slideNumber - 1 >= currentIndex);
                else
                    ShowSlide(slideNumber - 1);
            }
        }

        private void FindSlidesDialog()
        {
            string query = ViewerDialogs.Prompt(
                this,
                "Find in presentation",
                "Text to find:",
                "");

            if (string.IsNullOrWhiteSpace(query))
                return;

            List<int> matches = new List<int>();

            for (int i = 0; i < slideTexts.Count; i++)
            {
                string haystack =
                    (slideTitles.Count > i ? slideTitles[i] + "\n" : "") +
                    (slideTexts.Count > i ? slideTexts[i] : "") +
                    (speakerNotes.Count > i ? "\n" + speakerNotes[i] : "");

                if (haystack.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0)
                    matches.Add(i);
            }

            if (matches.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "No matching slides were found.",
                    "Find",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            ShowSlideListDialog("Find: " + query, matches);
        }

        private void ShowSlideSorter()
        {
            List<int> indexes = new List<int>();

            for (int i = 0; i < renderedSlides.Count; i++)
                indexes.Add(i);

            ShowSlideListDialog("All Slides", indexes);
        }

        private void ShowSlideListDialog(string title, List<int> indexes)
        {
            using (Form form = new Form())
            {
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.Width = 900;
                form.Height = 650;
                form.MinimumSize = new Size(680, 480);
                form.BackColor = ApplicationTheme.Window;
                form.ForeColor = ApplicationTheme.PrimaryText;

                FlowLayoutPanel flow = new FlowLayoutPanel();
                flow.Dock = DockStyle.Fill;
                flow.AutoScroll = true;
                flow.WrapContents = true;
                flow.Padding = new Padding(16);
                flow.BackColor = ApplicationTheme.Window;
                form.Controls.Add(flow);

                for (int m = 0; m < indexes.Count; m++)
                {
                    int index = indexes[m];

                    bool selected =
                        index == currentIndex;

                    Panel card = new Panel();
                    card.Width = 250;
                    card.Height = 180;
                    card.Margin = new Padding(8);
                    card.BackColor = selected
                        ? ApplicationTheme.AccentSoft
                        : ApplicationTheme.Surface;
                    card.Cursor = Cursors.Hand;

                    Panel accent = new Panel();
                    accent.Dock = DockStyle.Left;
                    accent.Width = 3;
                    accent.Visible = selected;
                    accent.BackColor = ApplicationTheme.Accent;
                    accent.Cursor = Cursors.Hand;

                    PictureBox pic = new PictureBox();
                    pic.Left = 10;
                    pic.Top = 10;
                    pic.Width = 230;
                    pic.Height = 130;
                    pic.SizeMode = PictureBoxSizeMode.Zoom;
                    pic.BackColor = Color.Black;
                    pic.Image = LoadImageUnlocked(renderedSlides[index]);
                    pic.Cursor = Cursors.Hand;

                    Label label = new Label();
                    label.Left = 10;
                    label.Top = 146;
                    label.Width = 230;
                    label.Height = 24;
                    label.AutoEllipsis = true;
                    label.BackColor = Color.Transparent;
                    label.ForeColor =
                        index < hiddenSlides.Count && hiddenSlides[index]
                        ? Color.FromArgb(126, 132, 142)
                        : selected
                            ? ApplicationTheme.PrimaryText
                            : ApplicationTheme.SecondaryText;
                    label.TextAlign = ContentAlignment.MiddleLeft;

                    string titleText =
                        index < slideTitles.Count && !string.IsNullOrWhiteSpace(slideTitles[index])
                        ? slideTitles[index]
                        : "Slide " + (index + 1).ToString();

                    label.Text = (index + 1).ToString() + ". " + titleText +
                        (index < hiddenSlides.Count && hiddenSlides[index] ? " [Hidden]" : "");

                    EventHandler click = delegate
                    {
                        ShowSlide(index);
                        form.DialogResult = DialogResult.OK;
                        form.Close();
                    };

                    EventHandler enter =
                        delegate
                        {
                            card.BackColor = selected
                                ? ApplicationTheme.AccentSoftHover
                                : ApplicationTheme.SurfaceHover;
                        };

                    EventHandler leave =
                        delegate
                        {
                            card.BackColor = selected
                                ? ApplicationTheme.AccentSoft
                                : ApplicationTheme.Surface;
                        };

                    card.Click += click;
                    pic.Click += click;
                    label.Click += click;
                    accent.Click += click;

                    card.MouseEnter += enter;
                    pic.MouseEnter += enter;
                    label.MouseEnter += enter;
                    accent.MouseEnter += enter;

                    card.MouseLeave += leave;
                    pic.MouseLeave += leave;
                    label.MouseLeave += leave;
                    accent.MouseLeave += leave;

                    card.Controls.Add(pic);
                    card.Controls.Add(label);
                    card.Controls.Add(accent);
                    if (selected)
                        accent.BringToFront();
                    flow.Controls.Add(card);
                }

                form.ShowDialog(this);

                foreach (Control control in flow.Controls)
                {
                    Panel panel = control as Panel;
                    if (panel == null) continue;

                    foreach (Control child in panel.Controls)
                    {
                        PictureBox pb = child as PictureBox;
                        if (pb != null && pb.Image != null)
                        {
                            Image img = pb.Image;
                            pb.Image = null;
                            img.Dispose();
                        }
                    }
                }
            }
        }

        private void ShowPresentationInfo()
        {
            int hiddenCount = 0;
            int noteCount = 0;

            for (int i = 0; i < hiddenSlides.Count; i++)
                if (hiddenSlides[i]) hiddenCount++;

            for (int i = 0; i < speakerNotes.Count; i++)
                if (!string.IsNullOrWhiteSpace(speakerNotes[i])) noteCount++;

            string message =
                "File: " + (string.IsNullOrEmpty(currentFile) ? "" : Path.GetFileName(currentFile)) + "\r\n" +
                "Slides: " + renderedSlides.Count.ToString() + "\r\n" +
                "Hidden slides: " + hiddenCount.ToString() + "\r\n" +
                "Slides with notes: " + noteCount.ToString() + "\r\n" +
                "Renderer: " + activeEngine;

            MessageBox.Show(
                this,
                message,
                "Presentation Info",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void TogglePresenterView()
        {
            if (presenterView != null && !presenterView.IsDisposed)
            {
                presenterView.Close();
                presenterView = null;
                return;
            }

            presenterView = new PresenterViewForm();

            presenterView.PreviousRequested += delegate
            {
                Previous();
            };

            presenterView.NextRequested += delegate
            {
                Next();
            };

            presenterView.GoToSlideRequested += delegate(int slideNumber)
            {
                int index = slideNumber - 1;

                if (index >= 0 && index < renderedSlides.Count)
                    ShowSlide(index);
            };

            presenterView.StartShowRequested += delegate
            {
                StartSlideshow(true);
            };

            presenterView.FormClosed += delegate
            {
                presenterView = null;
            };

            presenterView.Show(this);
            SyncPresenterView();
        }

        private void SyncPresenterView()
        {
            if (presenterView == null || presenterView.IsDisposed)
                return;

            Image current = null;
            Image next = null;

            try
            {
                if (currentIndex >= 0 && currentIndex < renderedSlides.Count)
                    current = LoadImageUnlocked(renderedSlides[currentIndex]);

                int nextIndex = currentIndex + 1;

                if (internalSlideShowMode)
                    nextIndex = FindVisibleSlide(nextIndex, 1);

                if (nextIndex >= 0 && nextIndex < renderedSlides.Count)
                    next = LoadImageUnlocked(renderedSlides[nextIndex]);

                string notes =
                    currentIndex >= 0 && currentIndex < speakerNotes.Count
                    ? speakerNotes[currentIndex]
                    : "";

                string title =
                    currentIndex >= 0 && currentIndex < slideTitles.Count
                    ? slideTitles[currentIndex]
                    : "";

                presenterView.UpdateSlide(
                    current,
                    next,
                    currentIndex + 1,
                    renderedSlides.Count,
                    title,
                    notes);
            }
            finally
            {
                if (current != null) current.Dispose();
                if (next != null) next.Dispose();
            }
        }

        private void PrintWithLayout()
        {
            if (renderedSlides.Count == 0)
                return;

            using (PrintLayoutDialog dialog = new PrintLayoutDialog())
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                PrintSlidesLayout(dialog.SelectedMode);
            }
        }
    }
}
