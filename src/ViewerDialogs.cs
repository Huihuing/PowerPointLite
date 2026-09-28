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
                form.Width = 430;
                form.Height = 160;

                Label label = new Label();
                label.Left = 12;
                label.Top = 14;
                label.Width = 390;
                label.Text = labelText;

                TextBox text = new TextBox();
                text.Left = 12;
                text.Top = 40;
                text.Width = 390;
                text.Text = initialValue ?? "";
                text.SelectAll();

                Button ok = new Button();
                ok.Text = "OK";
                ok.Left = 246;
                ok.Top = 76;
                ok.Width = 75;
                ok.DialogResult = DialogResult.OK;

                Button cancel = new Button();
                cancel.Text = "Cancel";
                cancel.Left = 327;
                cancel.Top = 76;
                cancel.Width = 75;
                cancel.DialogResult = DialogResult.Cancel;

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
