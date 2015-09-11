using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace albatross_desktop
{
    public partial class ConfigForm : Form
    {
        MainForm parent = null;
        public ConfigForm(MainForm p)
        {
            InitializeComponent();
            parent = p;
            excelpathtb.Text = Config.ReadIniKey("path", "excel", parent.iniFile);
            xmlpathtb.Text = Config.ReadIniKey("path", "xml", parent.iniFile);
        }

        private void cancelbtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void commitbtn_Click(object sender, EventArgs e)
        {
            //assignment
            Config.WriteIniKey("path", "excel", excelpathtb.Text, parent.iniFile);
            Config.WriteIniKey("path", "xml", xmlpathtb.Text, parent.iniFile);
            this.Close();
        }

        private void excelpathbtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbdlg = new FolderBrowserDialog();
            fbdlg.ShowDialog();
            excelpathtb.Text = fbdlg.SelectedPath;
        }

        private void xmlpathbtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbdlg = new FolderBrowserDialog();
            fbdlg.ShowDialog();
            xmlpathtb.Text = fbdlg.SelectedPath;
        }
    }
}
