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
internal enum PrintLayoutMode
    {
        FullSlide,
        NotesPage,
        Handout2,
        Handout4,
        Handout6
    }

internal sealed class PrintLayoutDialog : Form
    {
        private readonly ComboBox combo;

        public PrintLayoutMode SelectedMode
        {
            get
            {
                switch (combo.SelectedIndex)
                {
                    case 1: return PrintLayoutMode.NotesPage;
                    case 2: return PrintLayoutMode.Handout2;
                    case 3: return PrintLayoutMode.Handout4;
                    case 4: return PrintLayoutMode.Handout6;
                    default: return PrintLayoutMode.FullSlide;
                }
            }
        }

        public PrintLayoutDialog()
        {
            Text = "Print Layout";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Width = 390;
            Height = 180;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;

            Label label = new Label();
            label.Left = 16;
            label.Top = 18;
            label.Width = 340;
            label.Text = "Choose what to print:";
            Controls.Add(label);

            combo = new ComboBox();
            combo.Left = 16;
            combo.Top = 46;
            combo.Width = 340;
            combo.DropDownStyle = ComboBoxStyle.DropDownList;
            combo.Items.Add("Full page slides");
            combo.Items.Add("Notes pages");
            combo.Items.Add("Handouts - 2 slides per page");
            combo.Items.Add("Handouts - 4 slides per page");
            combo.Items.Add("Handouts - 6 slides per page");
            combo.SelectedIndex = 0;
            Controls.Add(combo);

            Button ok = new Button();
            ok.Text = "Print";
            ok.Left = 200;
            ok.Top = 86;
            ok.Width = 75;
            ok.DialogResult = DialogResult.OK;
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancel";
            cancel.Left = 281;
            cancel.Top = 86;
            cancel.Width = 75;
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }
    }
}
