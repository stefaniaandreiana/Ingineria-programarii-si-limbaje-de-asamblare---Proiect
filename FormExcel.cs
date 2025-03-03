using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectIP_2
{
    public partial class FormExcel : Form
    {
        public int percentage;
        public FormExcel()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = textBox1.Text;
            int value;
            if (int.TryParse(input, out value))
            {
                percentage = value;
                this.Close();
            }
            else
            {
                MessageBox.Show("Valoare invalida!");
            }
        }
    }
}
