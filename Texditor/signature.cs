using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Texditor
{
    public partial class signature : Form
    {
        public signature()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //  richTextBox1.WordWrap = true;

            try
            {
                //string currentdir = System.IO.Directory.GetCurrentDirectory().ToString();
                //richTextBox1.Text = currentdir;
                if (File.Exists("signature.txt"))
                {
                    File.WriteAllText("signature.txt", "");
                    if (richTextBox1.Text.Length > 0)
                        for (int i = 0; i < richTextBox1.Lines.Count(); i++)
                        {
                            string textbox = string.Format("textbox{0}", i);
                            File.AppendAllText("signature.txt", richTextBox1.Lines[i] + Environment.NewLine);
                        }
                    else
                    {
                        MessageBox.Show("Enter Signature");
                    }
                    MessageBox.Show("Signature Saved");
                    this.Close();
                }
                //MessageBox.Show("Yes It Exists"); }
                else
                {
                    File.Create("signature.txt");
                    //    MessageBox.Show("I have Created");
                    File.WriteAllText("signature.txt", "");
                    if (richTextBox1.Text.Length > 0)
                        for (int i = 0; i < richTextBox1.Lines.Count(); i++)
                        {
                            string textbox = string.Format("textbox{0}", i);
                            File.AppendAllText("signature.txt", richTextBox1.Lines[i] + Environment.NewLine);
                        }
                    else
                    {
                        MessageBox.Show("Enter Signature");
                    }
                    MessageBox.Show("Signature Saved");
                    this.Close();
                }
            }
            catch (Exception ee)
            { MessageBox.Show(ee.Message); }
        }

        private void signature_Load(object sender, EventArgs e)
        {
            if (File.Exists("signature.txt"))
            {
                string abc = File.ReadAllText((System.IO.Directory.GetCurrentDirectory() + "\\signature.txt"));
                richTextBox1.Text = "" + abc;
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
        }
    }
}