using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace albatross_desktop
{
    public partial class StartForm : Form
    {
        public StartForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MainForm mf = new MainForm();
            this.Hide();
            mf.WindowState = FormWindowState.Maximized;
            mf.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FileConvertForm fcf = new FileConvertForm(this);
            this.Hide();
            fcf.WindowState = FormWindowState.Maximized;
            fcf.ShowDialog();
            this.Show();
        }
    }
}
