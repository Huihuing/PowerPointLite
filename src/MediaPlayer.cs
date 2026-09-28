using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PptxViewer
{
    internal sealed class PortableMediaPlayerForm : Form
    {
        [DllImport(
            "winmm.dll",
            CharSet = CharSet.Auto,
            SetLastError = true)]
        private static extern int mciSendString(
            string command,
            StringBuilder returnValue,
            int returnLength,
            IntPtr callback);

        [DllImport(
            "winmm.dll",
            CharSet = CharSet.Auto)]
        private static extern bool mciGetErrorString(
            int errorCode,
            StringBuilder errorText,
            int errorTextSize);

        private readonly Panel videoPanel;
        private readonly Label titleLabel;
        private readonly Label timeLabel;
        private readonly TrackBar seekBar;
        private readonly Button playPauseButton;
        private readonly System.Windows.Forms.Timer positionTimer;

        private readonly string mediaPath;
        private readonly string alias;

        private bool opened;
        private bool paused;
        private int durationMs;
        private bool seekUpdating;

        private PortableMediaPlayerForm(string path)
        {
            mediaPath = path;
            alias =
                "pplt_" +
                Guid.NewGuid()
                    .ToString("N")
                    .Substring(0, 8);

            Text = "PowerPointLite Media";
            StartPosition = FormStartPosition.CenterParent;
            Width = 900;
            Height = 600;
            MinimumSize = new Size(520, 360);
            BackColor = Color.FromArgb(18, 20, 24);
            ForeColor = Color.WhiteSmoke;

            titleLabel = new Label();
            titleLabel.Dock = DockStyle.Top;
            titleLabel.Height = 34;
            titleLabel.TextAlign =
                ContentAlignment.MiddleLeft;
            titleLabel.Padding =
                new Padding(10, 0, 10, 0);
            titleLabel.Font =
                new Font(
                    "Segoe UI",
                    10f,
                    FontStyle.Bold);
            titleLabel.Text =
                Path.GetFileName(path);
            Controls.Add(titleLabel);

            Panel controls = new Panel();
            controls.Dock = DockStyle.Bottom;
            controls.Height = 72;
            controls.Padding =
                new Padding(8);
            controls.BackColor =
                Color.FromArgb(30, 33, 39);
            Controls.Add(controls);

            playPauseButton = new Button();
            playPauseButton.Text = "Pause";
            playPauseButton.Left = 8;
            playPauseButton.Top = 8;
            playPauseButton.Width = 76;
            playPauseButton.Height = 28;
            playPauseButton.Click += delegate
            {
                TogglePause();
            };
            controls.Controls.Add(playPauseButton);

            Button stopButton = new Button();
            stopButton.Text = "Stop";
            stopButton.Left = 90;
            stopButton.Top = 8;
            stopButton.Width = 70;
            stopButton.Height = 28;
            stopButton.Click += delegate
            {
                StopPlayback();
            };
            controls.Controls.Add(stopButton);

            Button replayButton = new Button();
            replayButton.Text = "Replay";
            replayButton.Left = 166;
            replayButton.Top = 8;
            replayButton.Width = 70;
            replayButton.Height = 28;
            replayButton.Click += delegate
            {
                Replay();
            };
            controls.Controls.Add(replayButton);

            timeLabel = new Label();
            timeLabel.Left = 250;
            timeLabel.Top = 12;
            timeLabel.Width = 180;
            timeLabel.Height = 22;
            timeLabel.Text =
                "00:00 / 00:00";
            controls.Controls.Add(timeLabel);

            seekBar = new TrackBar();
            seekBar.Left = 8;
            seekBar.Top = 40;
            seekBar.Width = 840;
            seekBar.Height = 26;
            seekBar.Anchor =
                AnchorStyles.Left |
                AnchorStyles.Right |
                AnchorStyles.Top;
            seekBar.Minimum = 0;
            seekBar.Maximum = 1000;
            seekBar.TickStyle =
                TickStyle.None;
            seekBar.Scroll += delegate
            {
                SeekFromBar();
            };
            controls.Controls.Add(seekBar);

            videoPanel = new Panel();
            videoPanel.Dock = DockStyle.Fill;
            videoPanel.BackColor =
                Color.Black;
            Controls.Add(videoPanel);
            videoPanel.BringToFront();

            positionTimer =
                new System.Windows.Forms.Timer();
            positionTimer.Interval = 250;
            positionTimer.Tick += delegate
            {
                UpdatePlaybackPosition();
            };

            videoPanel.Resize += delegate
            {
                ResizeMediaWindow();
            };

            Shown += delegate
            {
                ResizeMediaWindow();
            };

            FormClosed += delegate
            {
                positionTimer.Stop();
                CloseMedia();
            };
        }

        public static bool TryShow(
            IWin32Window owner,
            string path)
        {
            if (string.IsNullOrEmpty(path) ||
                !File.Exists(path))
            {
                return false;
            }

            using (
                PortableMediaPlayerForm form =
                    new PortableMediaPlayerForm(path))
            {
                string error;

                if (!form.TryOpenMedia(out error))
                {
                    CrashReporter.WriteLine(
                        "Internal media player failed: " +
                        error);
                    return false;
                }

                form.ShowDialog(owner);
                return true;
            }
        }

        private bool TryOpenMedia(
            out string error)
        {
            error = "";

            string quotedPath =
                "\"" + mediaPath.Replace(
                    "\"",
                    "\"\"") + "\"";

            int result =
                Send(
                    "open " +
                    quotedPath +
                    " alias " +
                    alias);

            if (result != 0)
            {
                result =
                    Send(
                        "open " +
                        quotedPath +
                        " type MPEGVideo alias " +
                        alias);
            }

            if (result != 0)
            {
                error =
                    GetMciErrorText(result);
                return false;
            }

            opened = true;

            Send(
                "set " +
                alias +
                " time format milliseconds");

            durationMs =
                QueryInt(
                    "status " +
                    alias +
                    " length");

            if (durationMs < 0)
                durationMs = 0;

            AttachVideoWindow();
            Send(
                "play " +
                alias);

            paused = false;
            playPauseButton.Text = "Pause";
            positionTimer.Start();
            UpdatePlaybackPosition();

            return true;
        }

        private void AttachVideoWindow()
        {
            if (!opened ||
                !videoPanel.IsHandleCreated)
            {
                return;
            }

            Send(
                "window " +
                alias +
                " handle " +
                videoPanel.Handle.ToInt64().ToString());

            ResizeMediaWindow();
        }

        private void ResizeMediaWindow()
        {
            if (!opened ||
                !videoPanel.IsHandleCreated)
            {
                return;
            }

            int width =
                Math.Max(
                    1,
                    videoPanel.ClientSize.Width);

            int height =
                Math.Max(
                    1,
                    videoPanel.ClientSize.Height);

            Send(
                "put " +
                alias +
                " window at 0 0 " +
                width.ToString() +
                " " +
                height.ToString());
        }

        private void TogglePause()
        {
            if (!opened)
                return;

            if (paused)
            {
                int result =
                    Send(
                        "resume " +
                        alias);

                if (result != 0)
                {
                    Send(
                        "play " +
                        alias);
                }

                paused = false;
                playPauseButton.Text =
                    "Pause";
            }
            else
            {
                Send(
                    "pause " +
                    alias);

                paused = true;
                playPauseButton.Text =
                    "Play";
            }
        }

        private void StopPlayback()
        {
            if (!opened)
                return;

            Send(
                "stop " +
                alias);

            paused = true;
            playPauseButton.Text =
                "Play";

            UpdatePlaybackPosition();
        }

        private void Replay()
        {
            if (!opened)
                return;

            Send(
                "seek " +
                alias +
                " to start");

            Send(
                "play " +
                alias);

            paused = false;
            playPauseButton.Text =
                "Pause";

            UpdatePlaybackPosition();
        }

        private void SeekFromBar()
        {
            if (!opened ||
                durationMs <= 0 ||
                seekUpdating)
            {
                return;
            }

            int target =
                (int)Math.Round(
                    durationMs *
                    (seekBar.Value / 1000.0));

            Send(
                "seek " +
                alias +
                " to " +
                target.ToString());

            if (!paused)
                Send(
                    "play " +
                    alias);

            UpdatePlaybackPosition();
        }

        private void UpdatePlaybackPosition()
        {
            if (!opened)
                return;

            int position =
                QueryInt(
                    "status " +
                    alias +
                    " position");

            if (position < 0)
                position = 0;

            if (durationMs <= 0)
            {
                durationMs =
                    QueryInt(
                        "status " +
                        alias +
                        " length");
            }

            seekUpdating = true;

            try
            {
                if (durationMs > 0)
                {
                    int value =
                        (int)Math.Round(
                            1000.0 *
                            position /
                            durationMs);

                    seekBar.Value =
                        Math.Max(
                            seekBar.Minimum,
                            Math.Min(
                                seekBar.Maximum,
                                value));
                }

                timeLabel.Text =
                    FormatTime(position) +
                    " / " +
                    FormatTime(durationMs);
            }
            finally
            {
                seekUpdating = false;
            }

            string mode =
                QueryString(
                    "status " +
                    alias +
                    " mode");

            if (string.Equals(
                    mode,
                    "stopped",
                    StringComparison.OrdinalIgnoreCase) &&
                durationMs > 0 &&
                position >= durationMs - 100)
            {
                paused = true;
                playPauseButton.Text =
                    "Replay";
            }
        }

        private void CloseMedia()
        {
            if (!opened)
                return;

            try
            {
                Send(
                    "stop " +
                    alias);
            }
            catch { }

            try
            {
                Send(
                    "close " +
                    alias);
            }
            catch { }

            opened = false;
        }

        private int Send(
            string command)
        {
            return mciSendString(
                command,
                null,
                0,
                IntPtr.Zero);
        }

        private string QueryString(
            string command)
        {
            StringBuilder buffer =
                new StringBuilder(512);

            int result =
                mciSendString(
                    command,
                    buffer,
                    buffer.Capacity,
                    IntPtr.Zero);

            if (result != 0)
                return "";

            return buffer
                .ToString()
                .Trim();
        }

        private int QueryInt(
            string command)
        {
            int value;
            string text =
                QueryString(command);

            return int.TryParse(
                text,
                out value)
                ? value
                : -1;
        }

        private static string GetMciErrorText(
            int errorCode)
        {
            StringBuilder buffer =
                new StringBuilder(512);

            if (mciGetErrorString(
                    errorCode,
                    buffer,
                    buffer.Capacity))
            {
                string text =
                    buffer.ToString().Trim();

                if (!string.IsNullOrEmpty(text))
                {
                    return
                        "MCI error " +
                        errorCode.ToString() +
                        ": " +
                        text;
                }
            }

            return
                "MCI error " +
                errorCode.ToString();
        }

        private static string FormatTime(
            int milliseconds)
        {
            if (milliseconds < 0)
                milliseconds = 0;

            TimeSpan time =
                TimeSpan.FromMilliseconds(
                    milliseconds);

            if (time.TotalHours >= 1.0)
            {
                return
                    ((int)time.TotalHours)
                    .ToString("00") +
                    ":" +
                    time.Minutes.ToString("00") +
                    ":" +
                    time.Seconds.ToString("00");
            }

            return
                time.Minutes.ToString("00") +
                ":" +
                time.Seconds.ToString("00");
        }
    }
}
