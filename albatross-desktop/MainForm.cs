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
        //DataSet ds = new DataSet();
        public string g_iniFile = System.IO.Directory.GetCurrentDirectory() + "\\conf.ini";
        Excel.Workbook g_wbb = null;
        Excel.Application g_eApp = null;
        string g_currFileName = null;
        string g_logFile = System.IO.Directory.GetCurrentDirectory() + "\\errlog.txt";
        FileStream g_fsFile = null;
        StreamWriter g_swWriter = null;
        StreamReader g_srReader = null;
        LogForm g_logform;

        #region "datagridview keyboard operations"
        DataTable g_dt = new DataTable();
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
            //Application.ApplicationExit += new EventHandler(Application_ApplicationExit);
            
            g_currFileName = @"d:\\type_copys_main.xml";
            processFile(g_currFileName);
        }

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
                readXml(fname);
            }
            else if (ext.Equals(".xls") || ext.Equals(".xlsx"))
            {
                readExcel(fname);
                //openExcel(fname);
            }
            return 0;
        }

        private int readXml(string fname)
        {
            XDocument doc = XDocument.Load(fname);
            bool colFinish = false;
            foreach (var item in doc.Root.Elements())
            {
                if (!colFinish)
                {
                    foreach (var attr in item.Attributes())
                    {
                        dgview.Columns.Add(attr.Name.ToString(), attr.Name.ToString());
                    }
                    colFinish = true;
                }
                int index = dgview.Rows.Add();
                int i = 0;
                foreach (var attr in item.Attributes())
                {
                    dgview.Rows[index].Cells[i].Value = attr.Value;
                    i++;
                }
            }

            return 0;
        }

        private int readXml2(string fname)
        {
            //ds.ReadXml(fname);
            //dgview.DataSource = ds.Tables[0];
            g_dt.Rows.Clear();
            g_dt.Columns.Clear();
            //return 0;
            XDocument doc = XDocument.Load(fname);
            bool colFinish = false;
            foreach (var item in doc.Root.Elements())
            {
                if (!colFinish)
                {
                    foreach (var attr in item.Attributes())
                    {
                        g_dt.Columns.Add(attr.Name.ToString(), System.Type.GetType("System.String"));
                    }
                    colFinish = true;
                }
                DataRow dr = g_dt.NewRow();
                //int index = dt.Rows.Add();
                int i = 0;
                foreach (var attr in item.Attributes())
                {
                    dr[attr.Name.ToString()] = attr.Value;
                    //dt.Rows[index].Cells[i].Value = attr.Value;
                    i++;
                }
                g_dt.Rows.Add(dr);
            }
            dgview.DataSource = g_dt;

            return 0;
        }

        private int readExcel(string fname)
        {
            dgview.Visible = true;
            panel1.Visible = true;
            panel1.Dock = DockStyle.Fill;
            dgview.Dock = DockStyle.Fill;
            string connStr = "";
            string fileType = System.IO.Path.GetExtension(fname);
            if (string.IsNullOrEmpty(fileType)) return -1;

            if (fileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fname + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fname + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "Select * FROM [{0}]";

            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dtSheetName = null;
            //DataSet ds = new DataSet();
            try
            {
                // 初始化连接，并打开
                conn = new OleDbConnection(connStr);
                conn.Open();
                // 获取数据源的表定义元数据                        
                string SheetName = "";
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                // 初始化适配器
                da = new OleDbDataAdapter();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];
                    if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                    {
                        continue;
                    }
                    da.SelectCommand = new OleDbCommand(String.Format(sql_F, SheetName), conn);
                    DataSet dsItem = new DataSet();
                    da.Fill(dsItem, SheetName);
                    //ds.Tables.Add(dsItem.Tables[0].Copy());
                    g_dt = dsItem.Tables[0].Copy();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                // 关闭连接
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
            //string a = ds.Tables[0].TableName;
            dgview.DataSource = g_dt;
            return 0;
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog savedlg = new SaveFileDialog();
            if (dgview.Rows.Count == 0)
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
                        saveXml(FileName);
                    }
                    else if (ext.Equals(".xls"))
                    {
                        saveExcel03(savedlg);
                    }
                    else if (ext.Equals(".xlsx"))
                    {
                        saveExcel03(FileName);
                    }

                }
            }
            return;
        }

        private void saveXml(string fname)
        {
            /*progressBar1.Visible = true;
            ds.WriteXml(fname);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                progressBar1.Value += 100 / ds.Tables[0].Rows.Count;
            }
            MessageBox.Show("数据已经成功导出到<" + fname, ">导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            progressBar1.Value = 0;
            progressBar1.Visible = false;*/
            //1、创建XML对象
            XDocument xdocument = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XComment("提示"));
            //2、创建跟节点
            XElement eRoot = new XElement("basenode");
            //添加到xdoc中
            xdocument.Add(eRoot);
            //3、添加子节点
            for (int j = 0; j < dgview.Rows.Count; j++)
            {
                XElement ele1 = new XElement("node");
                //ele1.Value = "内容1";
                eRoot.Add(ele1);
                for (int k = 0; k < dgview.Columns.Count; k++)
                {
                    //4、为ele1节点添加属性
                    XAttribute attr = new XAttribute(dgview.Columns[k].HeaderText, dgview.Rows[j].Cells[k].Value.ToString());
                    ele1.Add(attr);
                }
            }
            //5、快速添加子节点方法
            //eRoot.SetElementValue("子节点2", "内容2");
            //6、快速添加属性
            //ele1.SetAttributeValue("id", 12);
            //7、最后保存到文件，也可以写入到流中。
            xdocument.Save(fname);
        }

        private void saveExcel03(object obj)
        {
            SaveFileDialog saveFileDialog = obj as SaveFileDialog;
            Stream myStream;
            myStream = saveFileDialog.OpenFile();
            //StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding("gb2312"));
            StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
            string str = "";
            try
            {
                //写标题
                for (int i = 0; i < dgview.ColumnCount; i++)
                {
                    if (i > 0)
                    {
                        str += "\t";
                    }
                    str += dgview.Columns[i].HeaderText;
                }
                sw.WriteLine(str);
                //写内容
                for (int j = 0; j < dgview.Rows.Count; j++)
                {
                    string tempStr = "";
                    for (int k = 0; k < dgview.Columns.Count; k++)
                    {
                        if (k > 0)
                        {
                            tempStr += "\t";
                        }
                        tempStr += dgview.Rows[j].Cells[k].Value.ToString();
                    }

                    sw.WriteLine(tempStr);
                }
                sw.Close();
                myStream.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                sw.Close();
                myStream.Close();
            }
        }

        public bool saveExcel07(string fileName, bool isShowExcel = false)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                if (app == null)
                {
                    return false;
                }

                app.Visible = isShowExcel;
                Excel.Workbooks workbooks = app.Workbooks;
                Excel._Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Sheets sheets = workbook.Worksheets;
                Excel._Worksheet worksheet = (Excel._Worksheet)sheets.get_Item(1);
                if (worksheet == null)
                {
                    return false;
                }
                string sLen = "";
                //取得最后一列列名
                char H = (char)(64 + dgview.ColumnCount / 26);
                char L = (char)(64 + dgview.ColumnCount % 26);
                if (dgview.ColumnCount < 26)
                {
                    sLen = L.ToString();
                }
                else
                {
                    sLen = H.ToString() + L.ToString();
                }


                //标题
                string sTmp = sLen + "1";
                Excel.Range ranCaption = worksheet.get_Range(sTmp, "A1");
                string[] asCaption = new string[dgview.ColumnCount];
                for (int i = 0; i < dgview.ColumnCount; i++)
                {
                    asCaption[i] = dgview.Columns[i].HeaderText;
                }
                ranCaption.Value2 = asCaption;

                //数据
                object[] obj = new object[dgview.Columns.Count];
                for (int r = 0; r < dgview.RowCount - 1; r++)
                {
                    for (int l = 0; l < dgview.Columns.Count; l++)
                    {
                        if (dgview[l, r].ValueType == typeof(DateTime))
                        {
                            obj[l] = dgview[l, r].Value.ToString();
                        }
                        else
                        {
                            obj[l] = dgview[l, r].Value;
                        }
                    }
                    string cell1 = sLen + ((int)(r + 2)).ToString();
                    string cell2 = "A" + ((int)(r + 2)).ToString();
                    Excel.Range ran = worksheet.get_Range(cell1, cell2);
                    ran.Value2 = obj;
                }
                //保存
                workbook.SaveCopyAs(fileName);
                workbook.Saved = true;
            }
            finally
            {
                //关闭
                app.UserControl = false;
                app.Quit();
            }
            return true;
        }

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
                    int newindex = dgview.Rows.Add();
                    for (int icol = 0; icol < dgview.Columns.Count; icol++)
                    {
                        dgview.Rows[newindex].Cells[icol].Value = dgview.Rows[int.Parse(g_copiedRowIndices[i].ToString())].Cells[icol].Value;
                    }
                    dgview.Rows[newindex].Selected = true;
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
            saveXml(Config.ReadIniKey("path", "xml", g_iniFile) + "\\" + Path.GetFileNameWithoutExtension(g_currFileName) + ".xml");
            Log(LOGLEVEL.INFO, "按下了Control + Shift + X来保存表格数据到指定目录的Xml。");
        }

        private void saveEXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearLog();
            saveExcel07(Config.ReadIniKey("path", "excel", g_iniFile) + "\\" + Path.GetFileNameWithoutExtension(g_currFileName) + ".xlsx");
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

        /*private void Application_ApplicationExit(object sender, EventArgs e)
        {
            CloseExcelApplication();
            swWriter.Close();
        }*/
    }
}
