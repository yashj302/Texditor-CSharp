using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Windows.Documents;
using System.Windows.Controls;

namespace Texditor
{
    public partial class InsertTable : Form
    {
        Form1 frm = new Form1();
        public InsertTable(StringBuilder a)
        {
            
            InitializeComponent();
            richTextBox1.Rtf = "" + a;
            richTextBox1.Visible = false;

        }




        public void copyin(StringBuilder settingvalue)
        {

            StringBuilder replacevalue = settingvalue;


            frm.Whenreplaced(replacevalue);
            this.Hide();
            frm.ShowDialog();



        }




        private void InsertTable_Load(object sender, EventArgs e)
        {
            numericUpDown1.Minimum = 1;
            numericUpDown2.Minimum = 1;
            textBox1.Text = "1000";
            numericUpDown1.Maximum = 10;
            numericUpDown2.Maximum = 10;

        }





        public static String InsertTableInRichTextBox(int rows, int cols, int width)
        {
            //Create StringBuilder Instance
            StringBuilder sringTableRtf = new StringBuilder();

            //beginning of rich text format
            sringTableRtf.Append(@"{\rtf1 ");

            //Variable for cell width
            int cellWidth;

            //Start row
            sringTableRtf.Append(@"\trowd");

            //Loop to create table string
            for (int i = 0; i < rows; i++)
            {
                sringTableRtf.Append(@"\trowd");

                for (int j = 0; j < cols; j++)
                {
                    //Calculate cell end point for each cell
                    cellWidth = (j + 1) * width;

                    //A cell with width 1000 in each iteration.
                    sringTableRtf.Append(@"\cellx" + cellWidth.ToString());
                }

                //Append the row in StringBuilder
                sringTableRtf.Append(@"\intbl \cell \row");
            }
            sringTableRtf.Append(@"\pard");
            sringTableRtf.Append(@"}");

            return sringTableRtf.ToString();
        }


        private void button1_Click(object sender, EventArgs e)
        {




            try
            {
                if (textBox1.Text.Length > 0)
                {
                    StringBuilder repla = new StringBuilder(richTextBox1.Rtf);
                    int wdth = Convert.ToInt32(this.textBox1.Text);
                    int num = Convert.ToInt32(numericUpDown1.Value);
                    int num2 = Convert.ToInt32(numericUpDown2.Value);
                    richTextBox1.Visible = true;
                    richTextBox1.Rtf = InsertTableInRichTextBox(num, num2, wdth);
                    richTextBox1.Visible = false;
                    StringBuilder settin = new StringBuilder(richTextBox1.Rtf);
                    copyin(settin);
                }
                else
                {
                    MessageBox.Show("Enter Valid Row Width");
                }
                


            }
            catch (Exception Def)
            {
                MessageBox.Show(Def.Message);
            }






        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
                
        }
    }
}
