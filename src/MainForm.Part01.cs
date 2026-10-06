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
        private readonly Panel toolbar;
        private readonly SplitContainer mainSplit;
        private readonly FlowLayoutPanel thumbnails;
        private readonly Panel viewerPanel;
        private readonly PictureBox viewer;
        private readonly Label status;
        private readonly TrackBar zoomTrack;
        private readonly Label engineLabel;
        private readonly RichTextBox notesBox;
        private readonly ContextMenuStrip slideshowMenu;
        private readonly System.Windows.Forms.Timer sidebarHideTimer;

        private readonly List<string> renderedSlides = new List<string>();

        private PresenterViewForm presenterView;
        private bool autoHideSidebar;
        private bool sidebarMouseInside;
        private List<string> speakerNotes = new List<string>();
        private List<string> slideTexts = new List<string>();
        private List<string> slideTitles = new List<string>();
        private List<bool> hiddenSlides = new List<bool>();
        private List<List<InternalPptxRenderer.InteractiveRegion>> interactiveRegions =
            new List<List<InternalPptxRenderer.InteractiveRegion>>();
        private List<List<InternalPptxRenderer.ShapeRegion>> animationShapeRegions =
            new List<List<InternalPptxRenderer.ShapeRegion>>();
        private List<InternalPptxRenderer.TransitionSpec> transitionSpecs =
            new List<InternalPptxRenderer.TransitionSpec>();
        private List<List<string>> animationSteps = new List<List<string>>();
        private readonly ToolTip interactiveToolTip = new ToolTip();

        private string currentFile;
        private string currentCacheDirectory;
        private int currentIndex = -1;
        private int animationRevealCount;
        private bool internalSlideShowMode;
        private bool transitionAnimating;
        private bool blankScreenActive;
        private Color blankScreenColor = Color.Black;
        private Image blankScreenBackup;
        private string slideNumberBuffer = "";
        private float zoom = 1.0f;
        private bool fitMode = true;
        private bool fullscreen;
        private bool sidebarWasCollapsedBeforeFullscreen;
        private FormWindowState oldState;
        private FormBorderStyle oldBorder;

        private string activeEngine = "None";
        private string libreOfficePath;

        public MainForm(string startupFile)
        {
            CrashReporter.WriteLine("MainForm constructor entered");
            Text = "PowerPointLite 1.3";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 1400;
            Height = 850;
            MinimumSize = new Size(900, 600);
            KeyPreview = true;
            BackColor = Color.FromArgb(20, 22, 26);
            ForeColor = Color.WhiteSmoke;
            AllowDrop = true;

            toolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                BackColor = Color.FromArgb(30, 33, 39)
            };
            Controls.Add(toolbar);

            Button open = MakeButton("Open", 8, 8, 72);
            open.Click += delegate { OpenFile(); };

            Button prev = MakeButton("<", 86, 8, 44);
            prev.Click += delegate { Previous(); };

            Button next = MakeButton(">", 136, 8, 44);
            next.Click += delegate { Next(); };

            Button fit = MakeButton("Fit", 186, 8, 62);
            fit.Click += delegate { fitMode = true; ApplyFit(); };

            Button actual = MakeButton("100%", 254, 8, 62);
            actual.Click += delegate { fitMode = false; SetZoom(1.0f); };

            Button full = MakeButton("Full", 322, 8, 62);
            full.Click += delegate { ToggleFullscreen(); };

            Button show = MakeButton("F5 Show", 390, 8, 90);
            show.Click += delegate { StartSlideshow(false); };

            Button toc = MakeButton("TOC", 486, 8, 62);
            toc.Click += delegate { ToggleSidebar(); };

            Button print = MakeButton("Print", 554, 8, 68);
            print.Click += delegate { PrintWithLayout(); };

            Button recent = MakeButton("Recent", 628, 8, 72);
            recent.Click += delegate { ShowRecentFilesMenu(recent); };

            Button view = MakeButton("View", 706, 8, 62);
            view.Click += delegate { ShowViewMenu(view); };

            Button presenter = MakeButton("Presenter", 774, 8, 82);
            presenter.Click += delegate { TogglePresenterView(); };

            Button autoToc = MakeButton("Auto TOC", 862, 8, 82);
            autoToc.Click += delegate
            {
                autoHideSidebar = !autoHideSidebar;
                autoToc.Text = UiLocalization.Text(
                    autoHideSidebar ? "Auto TOC*" : "Auto TOC");
                LayoutViewerToolbar();

                if (autoHideSidebar && !fullscreen)
                {
                    mainSplit.Panel1Collapsed = true;
                    if (fitMode) BeginInvoke((MethodInvoker)ApplyFit);
                }
                else if (!autoHideSidebar && !fullscreen)
                {
                    mainSplit.Panel1Collapsed = false;
                    if (fitMode) BeginInvoke((MethodInvoker)ApplyFit);
                }
            };

            zoomTrack = new TrackBar
            {
                Minimum = 10,
                Maximum = 300,
                Value = 100,
                TickStyle = TickStyle.None,
                Width = 150,
                Height = 28,
                Left = 954,
                Top = 9
            };
            zoomTrack.Scroll += delegate
            {
                fitMode = false;
                SetZoom(zoomTrack.Value / 100.0f);
            };
            toolbar.Controls.Add(zoomTrack);

            engineLabel = new Label
            {
                AutoSize = false,
                Left = 1116,
                Top = 13,
                Width = 170,
                Height = 22,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                ForeColor = Color.Silver,
                Text = "Engine: detecting..."
            };
            toolbar.Controls.Add(engineLabel);

            status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 26,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(30, 33, 39),
                ForeColor = Color.Gainsboro,
                Text = "Drop a PPTX file here or click Open"
            };
            Controls.Add(status);

            // Use a real split layout.  In 0.7 the viewer's Fill panel could
            // still measure against the full form width while the TOC overlaid
            // the left side.  That made "Fit" too large and clipped the slide.
            mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                FixedPanel = FixedPanel.Panel1,
                SplitterWidth = 5,
                BackColor = Color.FromArgb(35, 38, 44)
            };
            Controls.Add(mainSplit);

            // Keep toolbar/status above the fill control in the docking order.
            toolbar.BringToFront();
            status.BringToFront();

            thumbnails = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(8),
                BackColor = Color.FromArgb(27, 30, 35)
            };
            mainSplit.Panel1.Controls.Add(thumbnails);

            notesBox = new RichTextBox
            {
                Dock = DockStyle.Bottom,
                Height = 150,
                Visible = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(245, 245, 242),
                ForeColor = Color.FromArgb(30, 30, 30),
                Font = new Font("Segoe UI", 10f),
                DetectUrls = true
            };
            notesBox.LinkClicked += delegate(object sender, LinkClickedEventArgs e)
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = e.LinkText,
                        UseShellExecute = true
                    });
                }
                catch { }
            };
            mainSplit.Panel2.Controls.Add(notesBox);

            viewerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.FromArgb(14, 15, 18)
            };
            mainSplit.Panel2.Controls.Add(viewerPanel);
            viewerPanel.BringToFront();

            viewer = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.Black
            };
            viewerPanel.Controls.Add(viewer);
            viewer.MouseClick += OnViewerMouseClick;
            viewer.MouseMove += OnViewerMouseMove;
            viewer.MouseLeave += delegate
            {
                viewer.Cursor = Cursors.Default;
                interactiveToolTip.SetToolTip(viewer, null);
            };

            slideshowMenu = new ContextMenuStrip();
            slideshowMenu.Items.Add("Next", null, delegate { Next(); });
            slideshowMenu.Items.Add("Previous", null, delegate { Previous(); });
            slideshowMenu.Items.Add(new ToolStripSeparator());
            slideshowMenu.Items.Add("Go to slide...", null, delegate { GoToSlideDialog(); });
            slideshowMenu.Items.Add("All slides...", null, delegate { ShowSlideSorter(); });
            slideshowMenu.Items.Add("Speaker notes", null, delegate { ToggleNotes(); });
            slideshowMenu.Items.Add(new ToolStripSeparator());
            slideshowMenu.Items.Add("Black screen", null, delegate { ToggleBlankScreen(Color.Black); });
            slideshowMenu.Items.Add("White screen", null, delegate { ToggleBlankScreen(Color.White); });
            slideshowMenu.Items.Add(new ToolStripSeparator());
            slideshowMenu.Items.Add("End show", null, delegate { EndInternalSlideshow(); });

            viewer.MouseUp += delegate(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Right && internalSlideShowMode)
                    slideshowMenu.Show(viewer, e.Location);
            };

            Resize += delegate
            {
                if (fitMode) BeginInvoke((MethodInvoker)ApplyFit);
            };

            mainSplit.SplitterMoved += delegate
            {
                if (fitMode) BeginInvoke((MethodInvoker)ApplyFit);
            };

            sidebarHideTimer = new System.Windows.Forms.Timer();
            sidebarHideTimer.Interval = 650;
            sidebarHideTimer.Tick += delegate
            {
                sidebarHideTimer.Stop();

                if (autoHideSidebar &&
                    !fullscreen &&
                    !sidebarMouseInside &&
                    !mainSplit.Panel1Collapsed)
                {
                    mainSplit.Panel1Collapsed = true;
                    if (fitMode) BeginInvoke((MethodInvoker)ApplyFit);
                }
            };

            thumbnails.MouseEnter += delegate
            {
                sidebarMouseInside = true;
                sidebarHideTimer.Stop();
            };

            thumbnails.MouseLeave += delegate
            {
                sidebarMouseInside = false;

                if (autoHideSidebar && !fullscreen)
                    sidebarHideTimer.Start();
            };

            mainSplit.MouseMove += delegate(object sender, MouseEventArgs e)
            {
                if (!autoHideSidebar || fullscreen)
                    return;

                if (mainSplit.Panel1Collapsed && e.X <= 18)
                {
                    mainSplit.Panel1Collapsed = false;
                    sidebarMouseInside = true;
                    sidebarHideTimer.Stop();

                    if (fitMode)
                        BeginInvoke((MethodInvoker)ApplyFit);
                }
                else if (!mainSplit.Panel1Collapsed && e.X > mainSplit.SplitterDistance + 40)
                {
                    sidebarMouseInside = false;
                    sidebarHideTimer.Stop();
                    sidebarHideTimer.Start();
                }
            };

            // Over the large slide area:
            //   wheel       = previous / next slide
            //   Ctrl+wheel  = zoom
            // Keep the thumbnail panel's wheel behavior as normal list scroll.
            viewerPanel.TabStop = true;

            viewerPanel.MouseEnter += delegate
            {
                viewerPanel.Focus();
            };

            viewer.MouseEnter += delegate
            {
                viewerPanel.Focus();
            };

            viewerPanel.MouseWheel += OnViewerMouseWheel;
            viewer.MouseWheel += OnViewerMouseWheel;

            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
            KeyDown += OnKeyDownMain;
            FormClosing += OnClosing;

            Shown += delegate
            {
                ConfigureInitialSplit();
                DetectBackends();

                if (!string.IsNullOrEmpty(startupFile) && File.Exists(startupFile))
                    BeginInvoke((MethodInvoker)delegate { LoadPresentation(startupFile); });
            };
        }

        private void ConfigureInitialSplit()
        {
            try
            {
                int width = Math.Max(600, mainSplit.ClientSize.Width);

                // Only set these after WinForms has performed the first layout.
                mainSplit.Panel1MinSize = 140;
                mainSplit.Panel2MinSize = 260;

                int desired = 260;
                int maximum = Math.Max(
                    mainSplit.Panel1MinSize,
                    width - mainSplit.Panel2MinSize - mainSplit.SplitterWidth);

                mainSplit.SplitterDistance = Math.Min(desired, maximum);
                CrashReporter.WriteLine(
                    "UI initialized. SplitterDistance=" + mainSplit.SplitterDistance.ToString());
            }
            catch (Exception ex)
            {
                // The viewer must still launch even on unusual DPI/layout setups.
                CrashReporter.WriteLine("Non-fatal splitter setup: " + ex.Message);
            }
        }

        private Button MakeButton(string text, int left, int top, int width)
        {
            Button b = new Button
            {
                Text = text,
                Left = left,
                Top = top,
                Width = width,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(42, 46, 54),
                ForeColor = Color.WhiteSmoke,
                TextAlign = ContentAlignment.MiddleCenter,
                UseCompatibleTextRendering = true,
                Padding = new Padding(0, 1, 0, 0),
                AutoSize = false,
                AutoEllipsis = true,
                Tag = text
            };
            b.FlatAppearance.BorderColor = Color.FromArgb(62, 67, 78);
            toolbar.Controls.Add(b);
            return b;
        }

        private void DetectBackends()
        {
            Type ppType = Type.GetTypeFromProgID("PowerPoint.Application");
            libreOfficePath = FindLibreOffice();

            if (ppType != null)
            {
                activeEngine = "Microsoft PowerPoint";
                engineLabel.Text = "Engine: PowerPoint Native";
            }
            else
            {
                activeEngine = "Internal OpenXML";
                engineLabel.Text = string.IsNullOrEmpty(libreOfficePath)
                    ? "Engine: Internal OpenXML"
                    : "Engine: Internal OpenXML + LibreOffice fallback";
            }
        }

        private static string FindLibreOffice()
        {
            List<string> candidates = new List<string>();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            candidates.Add(Path.Combine(baseDir, "LibreOffice", "program", "soffice.exe"));
            candidates.Add(Path.Combine(baseDir, "LibreOfficePortable", "App", "libreoffice", "program", "soffice.exe"));

            string pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (!string.IsNullOrEmpty(pf))
                candidates.Add(Path.Combine(pf, "LibreOffice", "program", "soffice.exe"));

            string pfx86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (!string.IsNullOrEmpty(pfx86))
                candidates.Add(Path.Combine(pfx86, "LibreOffice", "program", "soffice.exe"));

            foreach (string candidate in candidates)
            {
                try
                {
                    if (File.Exists(candidate))
                        return candidate;
                }
                catch { }
            }

            return null;
        }

        private void OpenFile()
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "PowerPoint|*.pptx;*.pptm;*.ppt|All files|*.*";
                dlg.Multiselect = false;

                if (dlg.ShowDialog(this) == DialogResult.OK)
                    LoadPresentation(dlg.FileName);
            }
        }
    }
}
