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
    public partial class Form1 : Form
    {
        Excel.Workbook wbb = null;
        Excel.Application eApp = null;
        WebBrowser webBrowser1 = null;
        public Form1()
        {
            InitializeComponent();
            
            openExcel("aaa");
        }

        private void openExcel(string sFileName)
        {
            webBrowser1 = new WebBrowser();
            string strFileName = @"d:\test.xlsx";
            Object refmissing = System.Reflection.Missing.Value;
            webBrowser1.Navigate(strFileName);
            object axWebBrowser = webBrowser1.ActiveXInstance;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser1.Visible = true;
            //webBrowser1.Dock = DockStyle.Fill;
            string ext = Path.GetExtension(e.Url.ToString());
            if (ext == ".xlsx" || ext == ".xls")
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
    }
}
