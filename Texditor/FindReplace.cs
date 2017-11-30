using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Texditor
{
    public partial class FindReplace : Form
    {
        Form1 frm = new Form1();
        public FindReplace(StringBuilder a)
        {
            InitializeComponent();
            
            richTextBox1.Rtf = ""+a;
        }

        private void FindReplace_Load(object sender, EventArgs e)
        {

        }
        
        public void copyin(StringBuilder settingvalue)
        {

            StringBuilder replacevalue = settingvalue;

            
            frm.Whenreplaced(replacevalue);
            this.Hide();
            frm.ShowDialog();
            
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox1.Text.Length > 0 && textBox2.Text.Length > 0)
                {
                    StringBuilder repla = new StringBuilder(richTextBox1.Rtf);
                    richTextBox1.Rtf = "" + repla.Replace(textBox1.Text + "", textBox2.Text + "");
                    StringBuilder settin = new StringBuilder(richTextBox1.Rtf);
                    copyin(settin);
                    
                }
                else {
                    MessageBox.Show("Enter Something");
                     }
                

            }
            catch (Exception Def)
            {
                MessageBox.Show(Def.Message);
            }

            }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder sb = new StringBuilder(richTextBox1.Rtf);
                copyin(sb);   
                
            }
            catch (Exception ee)
            { MessageBox.Show(ee.Message); }

            
        }
    }
}
