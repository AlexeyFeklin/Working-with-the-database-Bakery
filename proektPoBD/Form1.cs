using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace proektPoBD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TablesForm tablesForm = new TablesForm();
            tablesForm.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CheckForm checkForm = new CheckForm();
            checkForm.Show();
            this.Hide();
        }
    }
}
