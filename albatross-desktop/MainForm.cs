using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Net;
using System.Xml;
using Update;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Xml.Linq;
using System.Collections;
using System.Data.OleDb;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace albatross_desktop
{
    enum LOGLEVEL
    { 
        INFO = 1,
        WARN = 2,
        ERROR = 3,
    }
    public struct CellValueChangeLog
    {
        public string oldval;
        public string newval;
        public CellValueChangeLog(string o, string n)
        {
            oldval = o;
            newval = n;
        }
    }
    public partial class MainForm : Form
    {
        DataTable g_dt = null;
        public string g_iniFile = System.IO.Directory.GetCurrentDirectory() + "\\conf.ini";
        Excel.Workbook g_wbb = null;
        Excel.Application g_eApp = null;
        string g_currFileName = null;
        string g_logFile = System.IO.Directory.GetCurrentDirectory() + "\\errlog.txt";
        FileStream g_fsFile = null;
        StreamWriter g_swWriter = null;
        StreamReader g_srReader = null;
        LogForm g_logform;
        MySqlConnection mycon = null;

        #region "datagridview keyboard operations"
        ArrayList g_copiedRowIndices = new ArrayList();
        ArrayList g_copiedColIndices = new ArrayList();
        Dictionary<int, ArrayList> g_copiedCellIndices = new Dictionary<int,ArrayList>();
        #endregion
        #region "operation log"
        //Dictionary<int, ArrayList> cellValueChangeLog = new Dictionary<int, ArrayList>();//value member is CellValueChangeLog
        //ArrayList cellValueChangeLogList = new ArrayList();//member is cellValueChangeLog
        ArrayList g_operationLogList = new ArrayList();//member is cellValueChangeLogList
        int g_operationCurrentLogIndex = -1;
        string g_oldValue = "";
        #endregion

        public MainForm()
        {
            InitializeComponent();
            //checkUpdate();
            dgview.ShowCellToolTips = true;
            dgview.CellMouseEnter += new DataGridViewCellEventHandler(dgview_CellMouseEnter);
            g_fsFile = new FileStream(g_logFile, FileMode.OpenOrCreate);
            g_swWriter = new StreamWriter(g_fsFile);
            g_srReader = new StreamReader(g_fsFile);
            //this.WindowState = FormWindowState.Maximized;
            //Application.ApplicationExit += new EventHandler(Application_ApplicationExit);
            //open db
            /*try
            {
                string connstr = "server=" + Config.ReadIniKey("db", "host", g_iniFile) + ";";
                connstr += "User Id=" + Config.ReadIniKey("db", "user", g_iniFile) + ";";
                connstr += "password=" + Config.ReadIniKey("db", "pwd", g_iniFile) + ";";
                connstr += "Database=" + Config.ReadIniKey("db", "dbname", g_iniFile);
                mycon = new MySqlConnection(connstr);
                mycon.Open();
                //add sub menu
                ArrayList tmplist = new ArrayList();
                getDbData("show tables;", tmplist, null);
                foreach (ArrayList row in tmplist)
                {
                    foreach (string col in row)
                    {
                        dbDToolStripMenuItem.DropDownItems.Add(col, null, dbMenuClicked);
                    }
                }
            } 
            catch (Exception ex)
            {
                MessageBox.Show("db error: " + ex.ToString());
            }*/
            //g_currFileName = @"C:\Users\money_2\Desktop\data_round_client.txt.xls";
            //processFile(g_currFileName);
        }

        void dbMenuClicked(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel1.Dock = DockStyle.Fill;
            dgview.Dock = DockStyle.Fill;
            dgview.Visible = true;
            dgview.Rows.Clear();
            dgview.Columns.Clear();
            this.WindowState = FormWindowState.Maximized;
            ToolStripDropDownItem item = sender as ToolStripDropDownItem;
            //MessageBox.Show(item.Text);
            //get data from table
            ArrayList tmplist = new ArrayList();
            ArrayList tmpcolnamelist = new ArrayList();
            getDbData("select * from "+item.Text, tmplist, tmpcolnamelist);
            readDbData(tmplist, tmpcolnamelist);
        }

        #region "app updater"
        public void checkUpdate()
        {
            SoftUpdate app = new SoftUpdate(System.Windows.Forms.Application.ExecutablePath, "BlogWriter");
            app.UpdateFinish += new UpdateState(app_UpdateFinish);
            try
            {
                if (app.IsUpdate && MessageBox.Show("检查到新版本，是否更新？", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Thread update = new Thread(new ThreadStart(app.Update));
                    update.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void app_UpdateFinish()
        {
            MessageBox.Show("更新完成，请重新启动程序！", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        private void openBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog opendlg = new OpenFileDialog();
            opendlg.Filter = "xml文件(*.xml)|*.xml|excel2007文件(*.xlsx)|*.xlsx|所有文件(*.*)|*.*";

            if (opendlg.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opendlg.FileName;
                g_currFileName = sFileName;
                processFile(sFileName);
                this.WindowState = FormWindowState.Maximized;
            }
        }

        private int processFile(string fname)
        {
            string ext = Path.GetExtension(fname);
            string name = Path.GetFileName(fname);
            if (ext.Equals(".xml"))
            {
                panel1.Visible = true;
                panel1.Dock = DockStyle.Fill;
                dgview.Dock = DockStyle.Fill;
                dgview.Visible = true;
                g_dt = FileOpClass.readXml(fname);
                dgview.DataSource = null;
                dgview.DataSource = g_dt;
            }
            else if (ext.Equals(".xls") || ext.Equals(".xlsx"))
            {
                panel1.Visible = true;
                panel1.Dock = DockStyle.Fill;
                dgview.Dock = DockStyle.Fill;
                dgview.Visible = true;
                g_dt = FileOpClass.readExcel(fname);
                dgview.DataSource = null;
                dgview.DataSource = g_dt;
            }
            else if (ext.Equals(".txt"))
            {
                panel1.Visible = true;
                panel1.Dock = DockStyle.Fill;
                dgview.Dock = DockStyle.Fill;
                dgview.Visible = true;
                g_dt = FileOpClass.readJson(fname);
                dgview.DataSource = null;
                dgview.DataSource = g_dt;
            }
            return 0;
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog savedlg = new SaveFileDialog();
            if (g_dt != null && g_dt.Rows.Count == 0)
            {
                MessageBox.Show("没有数据可供导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                savedlg.Filter = "XML files (*.xml)|*.xml|EXCEL files (*.xlsx)|*.xlsx";
                savedlg.FilterIndex = 0;
                savedlg.RestoreDirectory = true;
                //saveFileDialog2.CreatePrompt = true; 
                savedlg.Title = "导出文件保存路径";
                savedlg.FileName = null;
                savedlg.ShowDialog();
                string FileName = savedlg.FileName;

                if (FileName.Length != 0)
                {
                    string ext = Path.GetExtension(FileName);
                    if (ext.Equals(".xml"))
                    {
                        FileOpClass.saveXml(FileName, g_dt);
                    }
                    else if (ext.Equals(".xls"))
                    {
                        FileOpClass.saveExcel03(savedlg, g_dt);
                    }
                    else if (ext.Equals(".xlsx"))
                    {
                        FileOpClass.saveExcel07(FileName, g_dt);
                    }
                }
            }
            return;
        }

        #region "webbrowser"
        private void openExcel(string sFileName)
        {
            webBrowser1.Visible = true;
            webBrowser1.Dock = DockStyle.Fill;
            //string strFileName = @"d:\test.xlsx";
            Object refmissing = System.Reflection.Missing.Value;
            webBrowser1.Navigate(sFileName);
            object axWebBrowser = webBrowser1.ActiveXInstance;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string ext = Path.GetExtension(e.Url.ToString());
            if (ext == ".xlsx" || ext == ".xls")
            {
                Object refmissing = System.Reflection.Missing.Value;
                object[] args = new object[4];
                //args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;
                args[0] = refmissing;
                args[1] = SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER;
                //此处SHDocVw需要添加此引用 c:/windows/system32/SHDocVw.dll
                args[2] = refmissing;
                args[3] = refmissing;
                object axWebBrowser = webBrowser1.ActiveXInstance;
                axWebBrowser.GetType().InvokeMember("ExecWB", BindingFlags.InvokeMethod, null, axWebBrowser, args);
                object oApplication = axWebBrowser.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, axWebBrowser, null);
                //此处BindingFlags需要添加 using System.Reflection;

                //wbb = (Excel.Workbook)oApplication;//wbb和eApp需要在一个全局变量里声明，以便可以回收

                g_eApp = ((Excel.Workbook)oApplication).Application;
                g_wbb = g_eApp.Workbooks[1];
                Excel.Worksheet ws = g_wbb.Worksheets[1] as Excel.Worksheet;
                ws.Cells.Font.Name = "Verdana";
                ws.Cells.Font.Size = 14;
                ws.Cells.Font.Bold = true;
                Excel.Range range = ws.Cells;
                Excel.Range oCell = range[10, 10] as Excel.Range;
                oCell.Value2 = "你好";
            }
            else
            {
                MessageBox.Show(e.Url.ToString());
            }
        }

        private void NAR(Object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }
        #endregion

        public void CloseExcelApplication()
        {
            try
            {
                /*for (int i = 0; i < wbb.Worksheets.Count; i++)
                {
                    NAR(wbb.Worksheets[i] as Excel.Worksheet);
                }*/
                if (null != g_wbb)
                {
                    g_wbb.Close(false);
                    NAR(g_wbb);
                }
                if (null != g_eApp)
                {
                    g_eApp.Quit();
                    NAR(g_eApp);
                }

            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            CloseExcelApplication();
            g_swWriter.Close();
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            //fixme!!! if multi files
            string fname = null;
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            fname = s[0];
            g_currFileName = fname;
            processFile(fname);
            this.WindowState = FormWindowState.Maximized;
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void AddRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dgview.Rows.Add();
        }

        private void ConfigToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ConfigForm confdlg = new ConfigForm(this);
            confdlg.ShowDialog();
        }

        private void Log(LOGLEVEL level, string text)
        {
            if (level == LOGLEVEL.WARN)
            {
                statusInfo.ForeColor = Color.Green;
            }
            else if (level == LOGLEVEL.ERROR)
            {
                statusInfo.ForeColor = Color.Red;
            }
            else
            {
                statusInfo.ForeColor = Color.Black;
            }
            statusInfo.Text = text;
            g_swWriter.WriteLine("[" + DateTime.Now.ToString() + "]" + text);
        }
        
        private void ClearLog()
        {
            statusInfo.ForeColor = Color.Black;
            statusInfo.Text = "";
        }
       
        private string ArrayListToStr(ArrayList list)
        {
            list.Sort();
            StringBuilder sb = new StringBuilder();
            foreach (int m in list)
            {
                sb.Append(m.ToString() + " ");
            }
            return sb.ToString();
        }
       
        private string DictToStr(Dictionary<int, ArrayList> dict)
        {
            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<int, ArrayList> kvp in dict)
            {
                sb.Append(kvp.Key + " : ");
                foreach (int m in kvp.Value)
                {
                    sb.Append(m.ToString() + ",");
                }
                sb.Append("|");
            }
            sb.Remove(sb.Length - 1, 1);
            return sb.ToString();
        }
        
        private bool isArrayListEquals(ArrayList a, ArrayList b)
        {
            if (a.Count != b.Count)
            {
                return false;
            }
            for (int i = 0; i < a.Count; i++)
            {
                if (!a[i].Equals(b[i]))
                {
                    return false;
                }
            }
            return true;
        }

        #region "DB"
        private void getDbData(string sql, ArrayList list, ArrayList colnamelist)
        {
            MySqlCommand mycmd = new MySqlCommand(sql, mycon);
            MySqlDataReader reader = mycmd.ExecuteReader();
            try
            {
                bool getcolnames = false;
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        ArrayList sublist = new ArrayList();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (reader.IsDBNull(i))
                            {
                                sublist.Add("");
                            }
                            else
                            {
                                sublist.Add(reader.GetString(i));
                            }
                            if (!getcolnames && colnamelist != null)
                            {
                                colnamelist.Add(reader.GetName(i));
                            }
                        }
                        list.Add(sublist);
                        getcolnames = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("查询失败: "+ex.ToString());
            }
            finally
            {
                reader.Close();
            }
        }

        private void readDbData(ArrayList datalist, ArrayList colnamelist)
        {
            foreach (var colname in colnamelist)
            {
                dgview.Columns.Add(colname.ToString(), colname.ToString());
            }
            foreach (var data in datalist)
            {
                int index = dgview.Rows.Add();
                int i = 0;
                ArrayList tmplist = data as ArrayList;
                foreach (var attr in tmplist)
                {
                    dgview.Rows[index].Cells[i].Value = attr.ToString();
                    i++;
                }
            }
        }
        #endregion

        private void copyOperation()
        {
            ClearLog();
            if (dgview.SelectedRows.Count > 0)
            {
                //copy rows
                g_copiedRowIndices.Clear();
                for (int i = 0; i < dgview.SelectedRows.Count; i++)
                {
                    g_copiedRowIndices.Add(dgview.SelectedRows[i].Index);
                }
                Log(LOGLEVEL.INFO, dgview.SelectedRows.Count.ToString() + "rows copied, row index in paste board: " + ArrayListToStr(g_copiedRowIndices));
            }
            else if (dgview.SelectedColumns.Count > 0)
            {
                //copy cols
                Log(LOGLEVEL.INFO, dgview.SelectedColumns.Count.ToString());
            }
            else if (dgview.SelectedCells.Count > 0)
            {
                //copy cells
                g_copiedCellIndices.Clear();
                for (int i = 0; i < dgview.SelectedCells.Count; i++)
                {
                    if (g_copiedCellIndices.ContainsKey(dgview.SelectedCells[i].RowIndex))
                    {
                        g_copiedCellIndices[dgview.SelectedCells[i].RowIndex].Add(dgview.SelectedCells[i].ColumnIndex);
                        g_copiedCellIndices[dgview.SelectedCells[i].RowIndex].Sort();
                    }
                    else
                    {
                        g_copiedCellIndices.Add(dgview.SelectedCells[i].RowIndex, new ArrayList() { dgview.SelectedCells[i].ColumnIndex });
                    }
                }
                Log(LOGLEVEL.INFO, dgview.SelectedCells.Count.ToString() + "cells copied, cell index in paste board: " + DictToStr(g_copiedCellIndices));
            }
            dgview.ClearSelection();
        }

        private void pasteOperation()
        {
            ClearLog();
            if (g_copiedRowIndices.Count > 0)
            {
                //paste copied rows
                for (int i = 0; i < g_copiedRowIndices.Count; i++)
                {
                    DataRow dr = g_dt.NewRow();
                    for (int icol = 0; icol < g_dt.Columns.Count; icol++)
                    {
                        dr[g_dt.Columns[icol].ColumnName] = g_dt.Rows[int.Parse(g_copiedRowIndices[i].ToString())][g_dt.Columns[icol].ColumnName].ToString();
                    }
                    g_dt.Rows.Add(dr);
                }
                Log(LOGLEVEL.INFO, g_copiedRowIndices.Count.ToString() + " rows pasted, row index in paste board: " + ArrayListToStr(g_copiedRowIndices));
            }
            else if (g_copiedColIndices.Count > 0)
            {
                //paste copied cols
            }
            else if (g_copiedCellIndices.Count > 0)
            {
                //paste copied cells
                //check if copied cells can be pasted
                int firstkey = 0;
                foreach (KeyValuePair<int, ArrayList> kvp in g_copiedCellIndices)
                {
                    if (0 == firstkey)
                    {
                        firstkey = kvp.Key;
                    }
                    if (!isArrayListEquals(g_copiedCellIndices[kvp.Key], g_copiedCellIndices[firstkey]))
                    {
                        Log(LOGLEVEL.ERROR, "cells cant be pasted, should be square box, cell index in paste board: " + DictToStr(g_copiedCellIndices));
                        return;
                    }
                }
                //check if select a postion
                if (dgview.SelectedCells.Count <= 0)
                {
                    Log(LOGLEVEL.ERROR, "should select paste position, cell index in paste board: " + DictToStr(g_copiedCellIndices));
                    return;
                }
                //check if paste position valid, out of bound
                int pasteStartRowDex = 999999;
                int pasteStartColDex = 999999;
                for (int i = 0; i < dgview.SelectedCells.Count; i++)
                {
                    //find first row index cell
                    if (pasteStartRowDex > dgview.SelectedCells[i].RowIndex)
                    {
                        pasteStartRowDex = dgview.SelectedCells[i].RowIndex;
                        pasteStartColDex = dgview.SelectedCells[i].ColumnIndex;
                    }
                }
                if (pasteStartRowDex + g_copiedCellIndices.Count > dgview.Rows.Count || pasteStartColDex + g_copiedCellIndices[firstkey].Count > dgview.Columns.Count)
                {
                    Log(LOGLEVEL.ERROR, "out of bound, cell index in paste board: " + DictToStr(g_copiedCellIndices));
                    return;
                }
                //start to paste cells
                ArrayList tmplist = new ArrayList();
                if (g_copiedCellIndices.Count == 1)
                {
                    foreach (KeyValuePair<int, ArrayList> kvp in g_copiedCellIndices)
                    {
                        if (kvp.Value.Count == 1)
                        {
                            //only one cell in source
                            for (int i = 0; i < dgview.SelectedCells.Count; i++)
                            {
                                foreach (int m in kvp.Value)
                                {
                                    writeToDictList(
                                        tmplist,
                                        dgview.SelectedCells[i].RowIndex,
                                        dgview.SelectedCells[i].ColumnIndex,
                                        dgview.SelectedCells[i].Value.ToString(),
                                        dgview.Rows[kvp.Key].Cells[m].Value.ToString()
                                    );
                                    dgview.SelectedCells[i].Value = dgview.Rows[kvp.Key].Cells[m].Value;
                                }
                            }
                            //operation log
                            writeSyncLog(tmplist);
                            return;
                        }
                    }
                }
                int tmprowdex = 0;
                foreach (KeyValuePair<int, ArrayList> kvp in g_copiedCellIndices)
                {
                    int tmpcoldex = 0;
                    foreach (int m in kvp.Value)
                    {
                        writeToDictList(
                            tmplist,
                            pasteStartRowDex + tmprowdex,
                            pasteStartColDex + tmpcoldex,
                            dgview.Rows[pasteStartRowDex + tmprowdex].Cells[pasteStartColDex + tmpcoldex].Value.ToString(),
                            dgview.Rows[kvp.Key].Cells[m].Value.ToString()
                        );
                        dgview.Rows[pasteStartRowDex + tmprowdex].Cells[pasteStartColDex + tmpcoldex].Value = dgview.Rows[kvp.Key].Cells[m].Value;
                        tmpcoldex++;
                    }
                    tmprowdex++;
                }
                writeSyncLog(tmplist);
            }
            dgview.DataSource = null;
            dgview.DataSource = g_dt;
        }

        private void undoOperation()
        {
            if (g_operationCurrentLogIndex < 0)
            {
                MessageBox.Show("no hist record!");
                return;
            }
            ArrayList valarrlist = null;
            if (g_operationCurrentLogIndex == 0)
            {
                valarrlist = g_operationLogList[0] as ArrayList;
                for (int i = 0; i < valarrlist.Count; i++)
                {
                    Dictionary<int, CellValueChangeLog> valdict = valarrlist[i] as Dictionary<int, CellValueChangeLog>;
                    foreach (KeyValuePair<int, CellValueChangeLog> kvp in valdict)
                    {
                        dgview.Rows[kvp.Key / 10000].Cells[kvp.Key % 10000].Value = kvp.Value.oldval;
                    }
                }
                return;
            }
            //fill data of index
            valarrlist = g_operationLogList[g_operationCurrentLogIndex] as ArrayList;
            for (int i = 0; i < valarrlist.Count; i++)
            {
                Dictionary<int, CellValueChangeLog> valdict = valarrlist[i] as Dictionary<int, CellValueChangeLog>;
                foreach (KeyValuePair<int, CellValueChangeLog> kvp in valdict)
                {
                    dgview.Rows[kvp.Key / 10000].Cells[kvp.Key % 10000].Value = kvp.Value.oldval;
                }
            }
            g_operationCurrentLogIndex--;
            dgview.ClearSelection();
            Log(LOGLEVEL.WARN, "undo: " + g_operationCurrentLogIndex.ToString() + "-" + g_operationLogList.Count.ToString());
        }

        private void redoOperation()
        {
            if (g_operationCurrentLogIndex < 0 || g_operationCurrentLogIndex > g_operationLogList.Count - 1)
            {
                MessageBox.Show("no more new record");
                return;
            }
            ArrayList valarrlist = null;
            if (g_operationCurrentLogIndex == g_operationLogList.Count - 1)
            {
                valarrlist = g_operationLogList[g_operationCurrentLogIndex] as ArrayList;
                for (int i = 0; i < valarrlist.Count; i++)
                {
                    Dictionary<int, CellValueChangeLog> valdict = valarrlist[i] as Dictionary<int, CellValueChangeLog>;
                    foreach (KeyValuePair<int, CellValueChangeLog> kvp in valdict)
                    {
                        dgview.Rows[kvp.Key / 10000].Cells[kvp.Key % 10000].Value = kvp.Value.newval;
                    }
                }
                return;
            }
            //fill data of index
            valarrlist = g_operationLogList[g_operationCurrentLogIndex] as ArrayList;
            for (int i = 0; i < valarrlist.Count; i++)
            {
                Dictionary<int, CellValueChangeLog> valdict = valarrlist[i] as Dictionary<int, CellValueChangeLog>;
                foreach (KeyValuePair<int, CellValueChangeLog> kvp in valdict)
                {
                    dgview.Rows[kvp.Key / 10000].Cells[kvp.Key % 10000].Value = kvp.Value.newval;
                }
            }
            g_operationCurrentLogIndex++;
            dgview.ClearSelection();
            Log(LOGLEVEL.WARN, "redo: " + g_operationCurrentLogIndex.ToString() + "-" + g_operationLogList.Count.ToString());
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Modifiers == Keys.Control && e.KeyCode == Keys.C) //copy
            {
                copyOperation();
            }
            else if (e.Modifiers == Keys.Control && e.KeyCode == Keys.V) //paste
            {
                pasteOperation();
            }
            else if (e.Modifiers == Keys.Control && e.KeyCode == Keys.Z)
            {
                undoOperation();
            }
            else if (e.Modifiers == Keys.Control && e.KeyCode == Keys.Y)
            {
                redoOperation();
            }
        }

        /*protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Right:
                    MessageBox.Show("Right");
                    break;
                case Keys.Left:
                    MessageBox.Show("Left");
                    break;
                case Keys.Up://方向键不反应
                    MessageBox.Show("up");
                    break;
                case Keys.Down:
                    MessageBox.Show("down");
                    break;
                case Keys.Space:
                    MessageBox.Show("space");
                    break;
                case Keys.Enter:
                    MessageBox.Show("enter");
                    break;
            }
            return false;//如果要调用KeyDown,这里一定要返回false才行,否则只响应重写方法里的按键.
            //这里调用一下父类方向,相当于调用普通的KeyDown事件.//所以按空格会弹出两个对话框
            //return base.ProcessCmdKey(ref msg, keyData);
        }*/

        void dgview_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0 || dgview.Rows.Count <= 0)
            {
                return;
            }
            dgview.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = dgview.Columns[e.ColumnIndex].HeaderText;
        }

        private void saveXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearLog();
            FileOpClass.saveXml(Config.ReadIniKey("path", "xml", g_iniFile) + "\\" + Path.GetFileNameWithoutExtension(g_currFileName) + ".xml", g_dt);
            Log(LOGLEVEL.INFO, "按下了Control + Shift + X来保存表格数据到指定目录的Xml。");
        }

        private void saveEXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearLog();
            FileOpClass.saveExcel07(Config.ReadIniKey("path", "excel", g_iniFile) + "\\" + Path.GetFileNameWithoutExtension(g_currFileName) + ".xlsx", g_dt);
            Log(LOGLEVEL.INFO, "按下了Control + Shift + E来保存表格数据到指定目录的Excel。");
        }

        private void dgview_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Log(LOGLEVEL.WARN, "begin edit " + e.RowIndex + " " + e.ColumnIndex + " " + dgview.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            g_oldValue = dgview.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
        }

        private void dgview_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Log(LOGLEVEL.WARN, "end edit " + e.RowIndex + " " + e.ColumnIndex + " " + dgview.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
            //operation log
            if (!g_oldValue.Equals(dgview.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()))
            {
                ArrayList tmplist = new ArrayList();
                writeToDictList(tmplist, e.RowIndex, e.ColumnIndex, g_oldValue, dgview.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                writeSyncLog(tmplist);
            }
        }

        private void dgview_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Log(LOGLEVEL.WARN, "value chnaged " + e.RowIndex + " " + e.ColumnIndex + " " + dgview.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
        }

        private void statusInfo_Click(object sender, EventArgs e)
        {
            g_logform = new LogForm(g_srReader);
            g_logform.Show();
        }

        private void writeSyncLog(ArrayList dictlist)
        {
            g_operationLogList.Add(dictlist);
            g_operationCurrentLogIndex = g_operationLogList.Count-1;
        }

        private void writeToDictList(ArrayList list, int rindex, int cindex, string oval, string nval)
        {
            Dictionary<int, CellValueChangeLog> celllog = new Dictionary<int, CellValueChangeLog>();
            int key = rindex*10000+cindex;
            celllog.Add(key, new CellValueChangeLog(oval, nval));
            list.Add(celllog);
        }

        private void convertVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileConvertForm fcf = new FileConvertForm(this);
            this.Hide();
            fcf.WindowState = FormWindowState.Maximized;
            fcf.ShowDialog();
            this.Show();
        }

        /*private void Application_ApplicationExit(object sender, EventArgs e)
        {
            CloseExcelApplication();
            swWriter.Close();
        }*/
    }
}
