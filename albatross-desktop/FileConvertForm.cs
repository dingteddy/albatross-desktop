using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace albatross_desktop
{
    public partial class FileConvertForm : Form
    {
        MainForm m_parent;
        string srcPath;
        string[] srcFiles = null;
        public FileConvertForm(MainForm parent)
        {
            InitializeComponent();

            m_parent = parent;
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            OpenFileDialog opendlg = new OpenFileDialog();
            opendlg.Filter = "xml文件(*.xml)|*.xml|excel2003文件(*.xls)|*.xls|excel2007文件(*.xlsx)|*.xlsx|JSON文件(*.txt)|*.txt";
            opendlg.Multiselect = true;

            if (opendlg.ShowDialog() == DialogResult.OK)
            {
                listBox1.Items.Clear();
                srcFiles = new string[opendlg.FileNames.Count()];
                int i = 0;
                foreach (string fname in opendlg.FileNames)
                {
                    srcPath = Path.GetDirectoryName(fname);
                    srcFiles[i] = Path.GetFileName(fname);
                    listBox1.Items.Add(srcFiles[i]);
                    i++;
                }
            }
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            int i = 0;
            foreach (string fname in s)
            {
                srcPath = Path.GetDirectoryName(fname);
                srcFiles[i] = Path.GetFileName(fname);
                listBox1.Items.Add(srcFiles[i]);
                i++;
            }
        }

        private void 打开OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBox1_DoubleClick(sender, e);
        }

        
    }
}
