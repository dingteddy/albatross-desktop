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
    public partial class ConfigForm : Form
    {
        public ConfigForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbdlg = new FolderBrowserDialog();
            fbdlg.ShowDialog();
            textBox1.Text = fbdlg.SelectedPath;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbdlg = new FolderBrowserDialog();
            fbdlg.ShowDialog();
            textBox2.Text = fbdlg.SelectedPath;
        }
    }
}
