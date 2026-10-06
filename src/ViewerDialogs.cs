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
internal static class ViewerDialogs
    {
        public static string Prompt(
            IWin32Window owner,
            string title,
            string labelText,
            string initialValue)
        {
            using (Form form = new Form())
            {
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowInTaskbar = false;
                form.ShowIcon = false;
                form.Width = 460;
                form.Height = 190;
                form.MinimumSize = new Size(420, 190);
                form.BackColor = ApplicationTheme.Window;
                form.ForeColor = ApplicationTheme.PrimaryText;
                form.Padding = new Padding(18);

                Label label = new Label();
                label.Left = 18;
                label.Top = 18;
                label.Width = 406;
                label.Height = 22;
                label.Text = labelText;
                label.ForeColor = ApplicationTheme.SecondaryText;
                label.AutoEllipsis = true;

                TextBox text = new TextBox();
                text.Left = 18;
                text.Top = 48;
                text.Width = 406;
                text.Height = 28;
                text.Text = initialValue ?? "";
                text.BackColor = ApplicationTheme.Surface;
                text.ForeColor = ApplicationTheme.PrimaryText;
                text.BorderStyle = BorderStyle.FixedSingle;
                text.Anchor =
                    AnchorStyles.Top |
                    AnchorStyles.Left |
                    AnchorStyles.Right;
                text.SelectAll();

                Button ok = new Button();
                ok.Text = "OK";
                ok.Left = 252;
                ok.Top = 96;
                ok.Width = 82;
                ok.Height = 32;
                ok.Anchor =
                    AnchorStyles.Bottom |
                    AnchorStyles.Right;
                ok.DialogResult = DialogResult.OK;
                ApplicationTheme.ApplyButton(ok);

                Button cancel = new Button();
                cancel.Text = "Cancel";
                cancel.Left = 342;
                cancel.Top = 96;
                cancel.Width = 82;
                cancel.Height = 32;
                cancel.Anchor =
                    AnchorStyles.Bottom |
                    AnchorStyles.Right;
                cancel.DialogResult = DialogResult.Cancel;
                ApplicationTheme.ApplyButton(cancel);

                form.Controls.Add(label);
                form.Controls.Add(text);
                form.Controls.Add(ok);
                form.Controls.Add(cancel);

                form.AcceptButton = ok;
                form.CancelButton = cancel;

                return form.ShowDialog(owner) == DialogResult.OK
                    ? text.Text
                    : null;
            }
        }
    }
}
