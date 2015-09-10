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

namespace albatross_desktop
{
    public partial class MainForm : Form
    {
        Excel.Workbook wbb = null;
        Excel.Application eApp = null;
        WebBrowser webBrowser1 = null;
        DataSet ds = new DataSet();

        public MainForm()
        {
            InitializeComponent();
            //checkUpdate();
            //processXml(@"d:\test.xml");
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
            opendlg.Filter = "xml文件(*.xml)|*.xml|所有文件(*.*)|*.*";

            if (opendlg.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opendlg.FileName;
                processFile(sFileName);
            }
        }

        private int processFile(string fname)
        {
            string ext = Path.GetExtension(fname);
            string name = Path.GetFileName(fname);
            if (ext.Equals(".xml"))
            {
                readXml(fname);
            }
            else if (ext.Equals(".xls") || ext.Equals(".xlsx"))
            {
                readExcel(fname);
            }
            return 0;
        }

        private int readXml(string fname)
        {
            ds.ReadXml(fname);
            dgview.DataSource = ds.Tables[0];
            //dgview.Rows.Clear();
            //dgview.Columns.Clear();
            dgview.Dock = DockStyle.Fill;
            dgview.Visible = true;
            //dgview.AllowUserToAddRows = true;
            return 0;
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
            //MessageBox.Show(dgview.Rows.Count.ToString());
            //MessageBox.Show(dgview.Rows[0].Cells[0].Value.ToString());

            return 0;
        }

        private int readExcel(string fname)
        {

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
            progressBar1.Visible = true;
            ds.WriteXml(fname);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                progressBar1.Value += 100 / ds.Tables[0].Rows.Count;
            }
            MessageBox.Show("数据已经成功导出到<" + fname, ">导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            progressBar1.Value = 0;
            progressBar1.Visible = false;
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
            webBrowser1 = new WebBrowser();
            webBrowser1.DocumentCompleted += webBrowser1_DocumentCompleted;
            //string strFileName = @"d:\test.xlsx";
            Object refmissing = System.Reflection.Missing.Value;
            webBrowser1.Navigate(sFileName);
            object axWebBrowser = webBrowser1.ActiveXInstance;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string ext = Path.GetExtension(e.Url.ToString());
            if (ext == "xlsx" || ext == "xls")
            {
                Object refmissing = System.Reflection.Missing.Value;
                object[] args = new object[4];
                args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;
                args[1] = SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER;
                //此处SHDocVw需要添加此引用 c:/windows/system32/SHDocVw.dll
                args[2] = refmissing;
                args[3] = refmissing;
                object axWebBrowser = webBrowser1.ActiveXInstance;
                axWebBrowser.GetType().InvokeMember("ExecWB", BindingFlags.InvokeMethod, null, axWebBrowser, args);
                object oApplication = axWebBrowser.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, axWebBrowser, null);
                //此处BindingFlags需要添加 using System.Reflection;

                //wbb = (Excel.Workbook)oApplication;//wbb和eApp需要在一个全局变量里声明，以便可以回收

                eApp = ((Excel.Workbook)oApplication).Application;
                wbb = eApp.Workbooks[1];
                Excel.Worksheet ws = wbb.Worksheets[1] as Excel.Worksheet;
                /*ws.Cells.Font.Name = "Verdana";
                ws.Cells.Font.Size = 14;
                ws.Cells.Font.Bold = true;
                Excel.Range range = ws.Cells;
                Excel.Range oCell = range[10, 10].asExcel.Range;
                oCell.Value2 = "你好";*/
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
                if (null != wbb)
                {
                    wbb.Close(false);
                    NAR(wbb);
                }
                if (null != eApp)
                {
                    eApp.Quit();
                    NAR(eApp);
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
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            //fixme!!! if multi files
            string fname = null;
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            fname = s[0];
            processFile(fname);
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
            MessageBox.Show("to do");
        }

        private void ConfigToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ConfigForm confdlg = new ConfigForm();
            confdlg.ShowDialog();
        }
    }
}
