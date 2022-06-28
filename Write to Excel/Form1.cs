using System;
using System.Windows.Forms;


namespace Write_to_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_Write_Click(object sender, EventArgs e)
        {
            WriteExcel.writeExcel();
            string a = textBox1.Text;
            string b = textBox2.Text;
            string c = textBox3.Text;
            string d = textBox4.Text;
            string[] things = new[] { a, b, c, d };
        }
    }
}
