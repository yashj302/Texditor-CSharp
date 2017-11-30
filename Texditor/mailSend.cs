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
    public partial class mailSend : Form
    {
        private StringBuilder richt;
        private string richtext;

        public mailSend(StringBuilder abc,string ab)
        {
            InitializeComponent();
            richtext = ab;
            richt = abc;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string from = textBox1.Text;
            string to = textBox3.Text;
            string password = textBox2.Text;
            string subject = richTextBox1.Text;
            Form1 frm = new Form1();
            frm.mailSSend(from, to, password, subject, richt, richtext);
            this.Close();
        }

        private void mailSend_Load(object sender, EventArgs e)
        {
        }
    }
}