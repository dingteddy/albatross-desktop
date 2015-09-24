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
        StartForm m_parent;
        string srcPath;
        string[] srcFiles = null;
        public FileConvertForm(StartForm parent)
        {
            InitializeComponent();

            m_parent = parent;

            string[] destFormatArr = { "xlsx", "xls", "xml", "txt" };
            foreach (string destFormat in destFormatArr)
            {
                targetFormatToolStripMenuItem.DropDownItems.Add(destFormat, null, targetFormatMenuClicked);
            }
            targetFormatMenuClicked(targetFormatToolStripMenuItem.DropDownItems[0], new EventArgs());
        }

        void targetFormatMenuClicked(object sender, EventArgs e)
        {
            ToolStripMenuItem senditem = sender as ToolStripMenuItem;
            foreach (ToolStripMenuItem item in targetFormatToolStripMenuItem.DropDownItems)
            {
                item.Checked = false;
            }
            senditem.Checked = true;
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
                srcFiles = new string[s.Count()];
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

        private void 转换RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string strFormat = "";
            foreach (ToolStripMenuItem item in targetFormatToolStripMenuItem.DropDownItems)
            {
                if (item.Checked)
                {
                    strFormat = item.Text;
                }
            }
            if (listBox1.Items.Count <= 0)
            {
                MessageBox.Show("should select source files first!");
                return;
            }
            MessageBox.Show("target format: " + strFormat);
            listBox1.SelectedIndex = -1;
            progressBar1.Visible = true;
            progressBar1.Maximum = listBox1.Items.Count;
            progressBar1.Value = 0;
            label1.Visible = true;
            label1.Location = new Point(progressBar1.Location.X+progressBar1.Width/2-label1.Width/2, progressBar1.Location.Y+progressBar1.Height/2-label1.Height/2);
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                progressBar1.Value += 1;
                label1.Text = progressBar1.Value + "/" + listBox1.Items.Count + " (" + listBox1.Items[i].ToString() + ")";
                processFile(srcPath + "\\" + listBox1.Items[i].ToString(), strFormat);
            }
            progressBar1.Value = listBox1.Items.Count;
            MessageBox.Show("all finished!");
            progressBar1.Value = 0;
            progressBar1.Visible = false;
            label1.Visible = false;
        }

        private int processFile(string fname, string target)
        {
            DataTable dt = null;
            string ext = Path.GetExtension(fname);
            string name = Path.GetFileName(fname);
            if (ext.Equals(".xml"))
            {
                dt = FileOpClass.readXml(fname);
            }
            else if (ext.Equals(".xls") || ext.Equals(".xlsx"))
            {
                dt = FileOpClass.readExcel(fname);
            }
            else if (ext.Equals(".txt"))
            {
                dt = FileOpClass.readJson(fname);
            }
            if (!Directory.Exists(srcPath + "\\target\\"))
            {
                Directory.CreateDirectory(srcPath + "\\target\\");
            }
            ////save
            if (dt == null)
            {
                return -1;
            }
            if (target.Equals("xml"))
            {
                FileOpClass.saveXml(srcPath + "\\target\\" + name + "." + target, dt);
            }
            else if (target.Equals("xls") || target.Equals("xlsx"))
            {
                FileOpClass.saveExcel07(srcPath + "\\target\\" + name + "." + target, dt);
            }
            else if (target.Equals("txt"))
            {
                //FileOpClass.save(srcPath + "\\target\\" + fname + "." + target, dt);
                MessageBox.Show("not supported!");
            }
            return 0;
        }

        private void 清空ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }
    }
}
