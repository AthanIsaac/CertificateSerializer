using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VisualSerializer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox3.Text, "[^0-9]"))
            {
                MessageBox.Show("Please enter only numbers.");
                textBox3.Text = textBox3.Text.Remove(textBox3.Text.Length - 1);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Enter valid source path");
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("Enter valid destination path");
            }
            else if (textBox3.Text == "")
            {
                MessageBox.Show("Enter a number of copies");
            }
            Program.CreateWordDocument(textBox1.Text.Trim(), textBox2.Text.Trim(), int.Parse(textBox3.Text));
            textBox3.Clear();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text.Replace('\"', ' ');
            textBox1.Text = textBox1.Text.Trim();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = textBox2.Text.Split('.')[0];
            textBox2.Text = textBox2.Text.Replace('\"', ' ');
            textBox2.Text = textBox2.Text.Trim();
            
        }
    }
}
