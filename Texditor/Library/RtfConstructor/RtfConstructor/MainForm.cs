using System;
using System.IO;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;

using ESCommon;
using ESCommon.Rtf;

namespace RtfConstructor
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
    
            Report report = new Report();
            RtfDocument rtf = report.GetRtf();
            RtfWriter rtfWriter = new RtfWriter();
            DateTime start;
            TimeSpan time;

            try
            {
                using (TextWriter writer = new StreamWriter("test.rtf"))
                {
                    start = DateTime.Now;

                    rtfWriter.Write(writer, rtf);

                    time = DateTime.Now - start;

                    label1.Text = String.Format("{0}.{1}", time.Seconds, time.Milliseconds.ToString().PadLeft(3, '0'));
                }
            }
            catch (IOException)
            {
                label1.Text = "I/O Exception";
            }

            button1.Enabled = true;
        }
    }
}