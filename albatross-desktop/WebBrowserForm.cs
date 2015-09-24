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
    public partial class WebBrowserForm : Form
    {
        Excel.Workbook g_wbb = null;
        Excel.Application g_eApp = null;
        public WebBrowserForm()
        {
            InitializeComponent();
        }

        private void wbf_DragDrop(object sender, DragEventArgs e)
        {
            //fixme!!! if multi files
            string fname = null;
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            fname = s[0];
            openExcel(fname);
            this.WindowState = FormWindowState.Maximized;
        }

        private void wbf_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        #region "webbrowser"
        private void openExcel(string sFileName)
        {
            webBrowser1.Visible = true;
            webBrowser1.Dock = DockStyle.Fill;
            //string strFileName = @"d:\test.xlsx";
            Object refmissing = System.Reflection.Missing.Value;
            webBrowser1.Navigate(sFileName);
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string ext = Path.GetExtension(e.Url.ToString());
            if (ext == ".xlsx" || ext == ".xls")
            {
                Object refmissing = System.Reflection.Missing.Value;
                object[] args = new object[4];
                args[0] = SHDocVw.OLECMDID.OLECMDID_HIDETOOLBARS;
                //args[0] = refmissing;
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
                ws.Cells.Font.Bold = false;
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
        }
    }
}
