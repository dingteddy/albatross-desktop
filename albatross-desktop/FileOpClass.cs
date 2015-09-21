using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;
using System.Data.OleDb;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace albatross_desktop
{
    class FileOpClass
    {
        //open
        public static DataTable readXml(string fname)
        {
            DataTable dt = new DataTable(Path.GetFileName(fname));
            XDocument doc = XDocument.Load(fname);
            bool colFinish = false;
            foreach (var item in doc.Root.Elements())
            {
                if (!colFinish)
                {
                    foreach (var attr in item.Attributes())
                    {
                        dt.Columns.Add(attr.Name.ToString());
                    }
                    colFinish = true;
                }
                DataRow dr = dt.NewRow();
                foreach (var attr in item.Attributes())
                {
                    dr[attr.Name.ToString()] = attr.Value;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public static DataTable readExcel(string fname)
        {
            DataTable dt = new DataTable(Path.GetFileName(fname));
            string connStr = "";
            string fileType = System.IO.Path.GetExtension(fname);
            if (string.IsNullOrEmpty(fileType))
            {
                return null;
            }

            if (fileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fname + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fname + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "Select * FROM [{0}]";

            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dtSheetName = null;
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
                    dt = dsItem.Tables[0].Copy();
                    break;
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
            return dt;
        }

        public static DataTable readJson(string fname)
        {
            /*string buffer = null;
            buffer = File.ReadAllText(fname);
            JObject jo = JsonConvert.DeserializeObject(buffer) as JObject;
            var s = from p in jo.Children()
                select p;
            foreach (var item in s)
            {
                //as a row
                JObject subjo = item as JObject;
                var subs = from p in subjo.Children()
                    select p;
                foreach (var subitem in s)
                {
                    //as a cell
                    MessageBox.Show(subitem.ToString());
                }
            }*/
            return null;
        }

        #region "save"
        public static void saveXml(string fname, DataTable dt)
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
                new XComment("this is comment region"));
            //2、创建跟节点
            XElement eRoot = new XElement("basenode");
            //添加到xdoc中
            xdocument.Add(eRoot);
            //3、添加子节点
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                XElement ele1 = new XElement("node");
                //ele1.Value = "内容1";
                eRoot.Add(ele1);
                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    //4、为ele1节点添加属性
                    XAttribute attr = new XAttribute(dt.Columns[k].ColumnName, dt.Rows[j][dt.Columns[k].ColumnName].ToString());
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

        public static void saveExcel03(object obj, DataTable dt)
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
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i > 0)
                    {
                        str += "\t";
                    }
                    str += dt.Columns[i].ColumnName;
                }
                sw.WriteLine(str);
                //写内容
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    string tempStr = "";
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        if (k > 0)
                        {
                            tempStr += "\t";
                        }
                        tempStr += dt.Rows[j][dt.Columns[k].ColumnName].ToString();
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

        public static bool saveExcel07(string fileName, DataTable dt, bool isShowExcel = false)
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
                char H = (char)(64 + dt.Columns.Count / 26);
                char L = (char)(64 + dt.Columns.Count % 26);
                if (dt.Columns.Count < 26)
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
                string[] asCaption = new string[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    asCaption[i] = dt.Columns[i].ColumnName;
                }
                ranCaption.Value2 = asCaption;

                //数据
                object[] obj = new object[dt.Columns.Count];
                for (int r = 0; r < dt.Rows.Count - 1; r++)
                {
                    for (int l = 0; l < dt.Columns.Count; l++)
                    {
                        //if (g_dt.Rows[r][g_dt.Columns[l].ColumnName].GetType() == typeof(DateTime))
                        //if (g_dt[l, r].ValueType == typeof(DateTime))
                        //{
                        //obj[l] = g_dt[l, r].Value.ToString();
                        obj[l] = dt.Rows[r][dt.Columns[l].ColumnName].ToString();
                        //}
                        //else
                        //{
                        //    obj[l] = g_dt.Rows[j][g_dt.Columns[k].ColumnName];
                        //}
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
        #endregion
    }
}
