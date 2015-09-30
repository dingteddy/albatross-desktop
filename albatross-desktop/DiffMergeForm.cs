using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace albatross_desktop
{
    public partial class DiffMergeForm : Form
    {
        public class celldex
        {
            public int row;
            public int col;
            public celldex(int i, int j)
            {
                row = i;
                col = j;
            }
        }

        string[] args = null;
        string g_destName = null;
        string g_srcName = null;
        DataTable g_destdt = null;
        DataTable g_srcdt = null;
        string g_iniFile = null;
        public DiffMergeForm(string[] args)
        {
            InitializeComponent();
            if (args != null && (args.Count() == 3 || args.Count() == 5))
            {
                /*for (int i = 0; i < args.Count(); i++)
                {
                    MessageBox.Show(args[i]);
                }*/
                this.args = args;
                if (args.Count() == 3)
                {
                    g_iniFile = args[2];
                    g_destName = args[0];
                    g_srcName = args[1];
                    processFile(srcdgview, g_srcName, ref g_srcdt);
                    processFile(destdgview, g_destName, ref g_destdt);
                    this.Text = g_destName + " | " + g_srcName;
                }
                else if (args.Count() == 5)
                {
                    g_iniFile = args[4];
                    g_srcName = args[1];
                    processFile(srcdgview, g_srcName, ref g_srcdt);
                    g_destName = args[0];
                    processFile(destdgview, g_destName, ref g_destdt);
                    DataTable tmpbasedt = null;
                    processFile(null, args[3], ref tmpbasedt);
                    this.Text = g_destName + " | " + g_srcName;
                }
                if (g_destdt != null && g_srcdt != null)
                {
                    compareFiles();
                }
            }
        }

        private void destdgview_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void destdgview_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            g_destName = s[0];
            processFile(destdgview, g_destName, ref g_destdt);
        }

        private void srcdgview_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void srcdgview_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            g_srcName = s[0];
            processFile(srcdgview, g_srcName, ref g_srcdt);
        }

        private int processFile(DataGridView viewer, string fname, ref DataTable dt)
        {
            //this.Text = fname;
            string ext = Path.GetExtension(fname);
            string name = Path.GetFileName(fname);
            Regex excelreg = new Regex(@".*xls[x]\.r\d*");
            if (excelreg.IsMatch(fname))
            {
                MessageBox.Show(fname + " matched!");
            }
            if (ext.Equals(".xml"))
            {
                dt = FileOpClass.readXml(fname);
                if (viewer != null)
                {
                    viewer.Dock = DockStyle.Fill;
                    viewer.Visible = true;
                    viewer.DataSource = null;
                    viewer.DataSource = dt;
                    FileOpClass.forbidGridViewSort(viewer, null);
                }
            }
            else if (ext.Equals(".xls") || ext.Equals(".xlsx") || excelreg.IsMatch(fname))
            {
                DataTable tmpdt = FileOpClass.readExcel(fname);
                ArrayList tiplist = new ArrayList();
                ArrayList tbtip = new ArrayList();
                MessageBox.Show(g_iniFile);
                if (!FileOpClass.rangeDt(tmpdt, ref dt, ref tiplist, ref tbtip, g_iniFile))
                {
                    return -1;
                }
                if (viewer != null)
                {
                    viewer.Dock = DockStyle.Fill;
                    viewer.Visible = true;
                    viewer.DataSource = null;
                    viewer.DataSource = dt;
                    FileOpClass.forbidGridViewSort(viewer, null);
                    //add tool tips
                    for (int i = 0; i < tiplist.Count; i++)
                    {
                        viewer.Columns[i].HeaderCell.ToolTipText = tiplist[i].ToString();//comment
                    }
                    if (tbtip.Count > 0)
                    {
                        viewer.TopLeftHeaderCell.ToolTipText = tbtip[0].ToString();//table comment
                    }
                }
            }
            else if (ext.Equals(".txt"))
            {
                dt = FileOpClass.readJson(fname);
                if (viewer != null)
                {
                    viewer.Dock = DockStyle.Fill;
                    viewer.Visible = true;
                    viewer.DataSource = null;
                    viewer.DataSource = dt;
                    FileOpClass.forbidGridViewSort(viewer, null);
                }
            }
            return 0;
        }

        private void destdgview_Scroll(object sender, ScrollEventArgs e)
        {
            srcdgview.FirstDisplayedScrollingRowIndex = destdgview.FirstDisplayedScrollingRowIndex;
            srcdgview.HorizontalScrollingOffset = destdgview.HorizontalScrollingOffset;
        }

        private void srcdgview_Scroll(object sender, ScrollEventArgs e)
        {
            destdgview.FirstDisplayedScrollingRowIndex = srcdgview.FirstDisplayedScrollingRowIndex;
            destdgview.HorizontalScrollingOffset = srcdgview.HorizontalScrollingOffset;
        }

        private void compareFiles()
        {
            ArrayList diffcells = new ArrayList();
            for (int irow = 0; irow < g_srcdt.Rows.Count; irow++)
            {
                for (int icol = 0; icol < g_srcdt.Columns.Count; icol++)
                {
                    if (irow < g_destdt.Rows.Count && icol < g_destdt.Columns.Count)
                    {
                        if (!g_srcdt.Rows[irow][g_srcdt.Columns[icol].ColumnName].Equals(g_destdt.Rows[irow][g_destdt.Columns[icol].ColumnName]))
                        {
                            diffcells.Add(new celldex(irow, icol));
                        }
                    }
                }
            }
            for (int i = 0; i < diffcells.Count; i++)
            {
                celldex tmp = diffcells[i] as celldex;
                srcdgview.Rows[tmp.row].Cells[tmp.col].Style.BackColor = Color.Green;
            }
        }

        private void 比较CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            compareFiles();
        }

        private void destdgview_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            srcdgview.ClearSelection();
            foreach (DataGridViewCell cell in destdgview.SelectedCells)
            {
                if (cell.RowIndex < srcdgview.Rows.Count-1 && cell.ColumnIndex < srcdgview.Columns.Count)
                {
                    srcdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Selected = true;
                }
            }
        }

        private void srcdgview_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            destdgview.ClearSelection();
            foreach (DataGridViewCell cell in srcdgview.SelectedCells)
            {
                if (cell.RowIndex < destdgview.Rows.Count - 1 && cell.ColumnIndex < destdgview.Columns.Count)
                {
                    destdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Selected = true;
                }
            }
        }

    }
}
