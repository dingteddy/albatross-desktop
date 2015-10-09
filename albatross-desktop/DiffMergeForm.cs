using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
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
        DataTable g_tmpbasedt = null;
        string g_iniFile = System.IO.Directory.GetCurrentDirectory() + "\\config.ini";
        public DiffMergeForm(string[] args)
        {
            InitializeComponent();

            string[] destFormatArr = { "按单元格", "按行", "按列" };
            foreach (string destFormat in destFormatArr)
            {
                比较模式ToolStripMenuItem.DropDownItems.Add(destFormat, null, targetFormatMenuClicked);
            }
            targetFormatMenuClicked(比较模式ToolStripMenuItem.DropDownItems[0], new EventArgs());

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
                    g_destName = args[1];//mine
                    g_srcName = args[0];
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
                    processFile(null, args[3], ref g_tmpbasedt);
                    this.Text = g_destName + " | " + g_srcName;
                    
                }
                
            }
            
        }

        void targetFormatMenuClicked(object sender, EventArgs e)
        {
            ToolStripMenuItem senditem = sender as ToolStripMenuItem;
            foreach (ToolStripMenuItem item in 比较模式ToolStripMenuItem.DropDownItems)
            {
                item.Checked = false;
            }
            senditem.Checked = true;
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
            /*if (excelreg.IsMatch(fname))
            {
                MessageBox.Show(fname + " matched!");
            }*/
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
                //MessageBox.Show(g_iniFile);
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

        private void compareFiles(DataGridView viewer, DataTable srcdt, DataTable destdt, Color markcolor)
        {
            ArrayList diffcells = new ArrayList();
            for (int irow = 0; irow < srcdt.Rows.Count; irow++)
            {
                for (int icol = 0; icol < srcdt.Columns.Count; icol++)
                {
                    if (irow < destdt.Rows.Count && icol < destdt.Columns.Count)
                    {
                        if (!srcdt.Rows[irow][srcdt.Columns[icol].ColumnName].Equals(destdt.Rows[irow][destdt.Columns[icol].ColumnName]))
                        {
                            diffcells.Add(new celldex(irow, icol));
                        }
                    }
                }
            }
            for (int i = 0; i < diffcells.Count; i++)
            {
                celldex tmp = diffcells[i] as celldex;
                viewer.Rows[tmp.row].Cells[tmp.col].Style.BackColor = markcolor;
            }
        }

        private void compareFilesByRow(DataGridView viewer, DataTable srcdt, DataTable destdt, Color markcolor)
        {
            ArrayList diffrows = new ArrayList();
            for (int irow = 0; irow < srcdt.Rows.Count; irow++)
            {
                    if (irow < destdt.Rows.Count)
                    {
                        DataRow dr1 = srcdt.Rows[irow];
                        DataRow dr2 = destdt.Rows[irow];
                        if (dr1.ItemArray.Count() != dr2.ItemArray.Count())
                        {
                            diffrows.Add(irow);
                            goto EXIT;
                        }
                        for (int i = 0; i < dr1.ItemArray.Count(); i++)
                        {
                            if (!dr1.ItemArray[i].Equals(dr2.ItemArray[i]))
                            {
                                diffrows.Add(irow);
                                goto EXIT;
                            }
                        }
                    }
            }
EXIT:
            for (int i = 0; i < diffrows.Count; i++)
            {
                for (int j = 0; j < viewer.Columns.Count; j++)
                {
                    viewer.Rows[int.Parse(diffrows[i].ToString())].Cells[j].Style.BackColor = markcolor;
                }
            }
        }

        private void compareFilesByCol(DataGridView viewer, DataTable srcdt, DataTable destdt, Color markcolor)
        {
            ArrayList diffrows = new ArrayList();
            for (int irow = 0; irow < srcdt.Rows.Count; irow++)
            {
                if (irow < destdt.Rows.Count)
                {
                    DataRow dr1 = srcdt.Rows[irow];
                    DataRow dr2 = destdt.Rows[irow];
                    if (dr1.ItemArray.Count() != dr2.ItemArray.Count())
                    {
                        diffrows.Add(irow);
                        goto EXIT;
                    }
                    for (int i = 0; i < dr1.ItemArray.Count(); i++)
                    {
                        if (!dr1.ItemArray[i].Equals(dr2.ItemArray[i]))
                        {
                            diffrows.Add(irow);
                            goto EXIT;
                        }
                    }
                }
            }
        EXIT:
            for (int i = 0; i < diffrows.Count; i++)
            {
                for (int j = 0; j < viewer.Columns.Count; j++)
                {
                    viewer.Rows[int.Parse(diffrows[i].ToString())].Cells[j].Style.BackColor = markcolor;
                }
            }
        }

        private void 比较CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string strFormat = "";
            foreach (ToolStripMenuItem item in 比较模式ToolStripMenuItem.DropDownItems)
            {
                if (item.Checked)
                {
                    strFormat = item.Text;
                }
            }
            if (this.args == null)
            {
                if (g_destdt != null && g_srcdt != null)
                {
                    destdgview.DataSource = null;
                    destdgview.DataSource = g_destdt;
                    FileOpClass.forbidGridViewSort(destdgview, null);
                    //compareFiles(destdgview, g_srcdt, g_destdt, Color.Green);
                    compareRaw(strFormat, destdgview, g_srcdt, g_destdt, Color.Green);
                }
                return;
            }
            if (this.args.Count() == 3)
            {
                if (g_destdt != null && g_srcdt != null)
                {
                    //compareFiles(destdgview, g_srcdt, g_destdt, Color.Green);
                    compareRaw(strFormat, destdgview, g_srcdt, g_destdt, Color.Green);
                }
            }
            else if (args.Count() == 5)
            {
                if (g_destdt != null && g_srcdt != null && g_tmpbasedt != null)
                {
                    //compareFiles(destdgview, g_tmpbasedt, g_destdt, Color.Gray);
                    compareRaw(strFormat, destdgview, g_tmpbasedt, g_destdt, Color.Gray);
                    //compareFiles(srcdgview, g_tmpbasedt, g_srcdt, Color.Red);
                    compareRaw(strFormat, srcdgview, g_tmpbasedt, g_srcdt, Color.Red);
                }
            }
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

        private void 使用左边ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (destdgview.SelectedCells.Count <= 0)
            {
                return;
            }
            foreach (DataGridViewCell cell in destdgview.SelectedCells)
            {
                if (cell.RowIndex < srcdgview.Rows.Count - 1 && cell.ColumnIndex < srcdgview.Columns.Count)
                {
                    g_srcdt.Rows[cell.RowIndex][g_srcdt.Columns[cell.ColumnIndex].ColumnName] = destdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value;
                    srcdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value = destdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value;
                    srcdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Style.BackColor = Color.White;
                    destdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Style.BackColor = Color.White;
                }
            }
            srcdgview.DataSource = null;
            srcdgview.DataSource = g_srcdt;
            srcdgview.ClearSelection();
            destdgview.ClearSelection();
        }

        private void 使用右边ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (srcdgview.SelectedCells.Count <= 0)
            {
                return;
            }
            //MessageBox.Show(srcdgview.SelectedCells.Count.ToString());
            foreach (DataGridViewCell cell in srcdgview.SelectedCells)
            {
                //MessageBox.Show(cell.RowIndex.ToString() + "--" + cell.ColumnIndex.ToString());
                if (cell.RowIndex < destdgview.Rows.Count - 1 && cell.ColumnIndex < destdgview.Columns.Count)
                {
                    destdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value = srcdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value;
                    srcdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Style.BackColor = Color.White;
                    destdgview.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Style.BackColor = Color.White;
                }
            }
            srcdgview.ClearSelection();
            destdgview.ClearSelection();
        }

        private void DiffMergeForm_Activated(object sender, EventArgs e)
        {
            if (this.args != null)
            {
                比较CToolStripMenuItem_Click(sender, e);
            }
        }

        public void UploadMultipart(byte[] file, string filename, string contentType, string url)
        {
            var webClient = new WebClient();
            string boundary = "------------------------" + DateTime.Now.Ticks.ToString("x");
            webClient.Headers.Add("Content-Type", "multipart/form-data; boundary=" + boundary);
            var fileData = webClient.Encoding.GetString(file);
            var package = string.Format("--{0}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n{3}\r\n--{0}--\r\n", boundary, filename, contentType, fileData);

            var nfile = webClient.Encoding.GetBytes(package);

            byte[] resp = webClient.UploadData(url, "POST", nfile);
            MessageBox.Show(resp.ToString());
        }

        /// <summary>
        /// Creates HTTP POST request & uploads database to server. Author : Farhan Ghumra
        /// </summary>
        private void UploadFilesToServer(Uri uri, Dictionary<string, string> data, string fileName, string fileContentType, byte[] fileData)
        {
            string boundary = "----------" + DateTime.Now.Ticks.ToString("x");
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpWebRequest.ContentType = "multipart/form-data; boundary=" + boundary;
            httpWebRequest.Method = "POST";
            httpWebRequest.BeginGetRequestStream((result) =>
            {
                try
                {
                    HttpWebRequest request = (HttpWebRequest)result.AsyncState;
                    using (Stream requestStream = request.EndGetRequestStream(result))
                    {
                        WriteMultipartForm(requestStream, boundary, data, fileName, fileContentType, fileData);
                    }
                    request.BeginGetResponse(a =>
                    {
                        try
                        {
                            var response = request.EndGetResponse(a);
                            var responseStream = response.GetResponseStream();
                            using (var sr = new StreamReader(responseStream))
                            {
                                using (StreamReader streamReader = new StreamReader(response.GetResponseStream()))
                                {
                                    string responseString = streamReader.ReadToEnd();
                                    //responseString is depend upon your web service.
                                    if (responseString == "Success")
                                    {
                                        MessageBox.Show("Backup stored successfully on server.");
                                    }
                                    else
                                    {
                                        MessageBox.Show("Error occurred while uploading backup on server.");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }, httpWebRequest);
        }

        /// <summary>
        /// Writes multi part HTTP POST request. Author : Farhan Ghumra
        /// </summary>
        private void WriteMultipartForm(Stream s, string boundary, Dictionary<string, string> data, string fileName, string fileContentType, byte[] fileData)
        {
            /// The first boundary
            byte[] boundarybytes = Encoding.UTF8.GetBytes("--" + boundary + "\r\n");
            /// the last boundary.
            byte[] trailer = Encoding.UTF8.GetBytes("\r\n--" + boundary + "–-\r\n");
            /// the form data, properly formatted
            string formdataTemplate = "Content-Dis-data; name=\"{0}\"\r\n\r\n{1}";
            /// the form-data file upload, properly formatted
            string fileheaderTemplate = "Content-Dis-data; name=\"{0}\"; filename=\"{1}\";\r\nContent-Type: {2}\r\n\r\n";

            /// Added to track if we need a CRLF or not.
            bool bNeedsCRLF = false;

            if (data != null)
            {
                foreach (string key in data.Keys)
                {
                    /// if we need to drop a CRLF, do that.
                    if (bNeedsCRLF)
                        WriteToStream(s, "\r\n");

                    /// Write the boundary.
                    WriteToStream(s, boundarybytes);

                    /// Write the key.
                    WriteToStream(s, string.Format(formdataTemplate, key, data[key]));
                    bNeedsCRLF = true;
                }
            }

            /// If we don't have keys, we don't need a crlf.
            if (bNeedsCRLF)
                WriteToStream(s, "\r\n");

            WriteToStream(s, boundarybytes);
            WriteToStream(s, string.Format(fileheaderTemplate, "file", fileName, fileContentType));
            /// Write the file data to the stream.
            WriteToStream(s, fileData);
            WriteToStream(s, trailer);
        }

        /// <summary>
        /// Writes string to stream. Author : Farhan Ghumra
        /// </summary>
        private void WriteToStream(Stream s, string txt)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(txt);
            s.Write(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// Writes byte array to stream. Author : Farhan Ghumra
        /// </summary>
        private void WriteToStream(Stream s, byte[] bytes)
        {
            s.Write(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// 向指定的URL地址发起一个POST请求，同时可以上传一些数据项以及上传文件。
        /// </summary>
        /// <param name="url">要请求的URL地址</param>
        /// <param name="keyvalues">要上传的数据项</param>
        /// <param name="fileList">要上传的文件列表</param>
        /// <param name="encoding">发送数据项，接收的字符编码方式</param>
        /// <returns>服务器的返回结果</returns>
        public static string SendHttpRequestPost(string url, Dictionary<string, string> keyvalues, Dictionary<string, string> fileList, Encoding encoding)
        {
            //if (fileList == null)
               // return SendHttpRequestPost(url, keyvalues, encoding);

            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException("url");

            if (encoding == null)
                encoding = Encoding.UTF8;


            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";        // 要上传文件，一定要是POST方法

            // 数据块的分隔标记，用于设置请求头，注意：这个地方最好不要使用汉字。
            string boundary = "---------------------------" + Guid.NewGuid().ToString("N");
            // 数据块的分隔标记，用于写入请求体。
            //   注意：前面多了一段： "--" ，而且它们将独占一行。
            byte[] boundaryBytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

            // 设置请求头。指示是一个上传表单，以及各数据块的分隔标记。
            request.ContentType = "multipart/form-data; boundary=" + boundary;

            // 先得到请求流，准备写入数据。
            Stream stream = request.GetRequestStream();

            if (keyvalues != null && keyvalues.Count > 0)
            {
                // 写入非文件的keyvalues部分
                foreach (KeyValuePair<string, string> kvp in keyvalues)
                {
                    // 写入数据块的分隔标记
                    stream.Write(boundaryBytes, 0, boundaryBytes.Length);

                    // 写入数据项描述，这里的Value部分可以不用URL编码
                    string str = string.Format(
                            "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}",
                            kvp.Key, kvp.Value);

                    byte[] data = encoding.GetBytes(str);
                    stream.Write(data, 0, data.Length);
                }
            }

            // 写入要上传的文件
            foreach (KeyValuePair<string, string> kvp in fileList)
            {
                // 写入数据块的分隔标记
                stream.Write(boundaryBytes, 0, boundaryBytes.Length);

                // 写入文件描述，这里设置一个通用的类型描述：application/octet-stream，具体的描述在注册表里有。
                string description = string.Format(
                        "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\n" +
                        "Content-Type: application/octet-stream\r\n\r\n",
                        kvp.Key, Path.GetFileName(kvp.Value));

                // 注意：这里如果不使用UTF-8，对于汉字会有乱码。
                byte[] header = Encoding.UTF8.GetBytes(description);
                stream.Write(header, 0, header.Length);

                // 写入文件内容
                byte[] body = File.ReadAllBytes(kvp.Value);
                stream.Write(body, 0, body.Length);
            }


            // 写入结束标记
            boundaryBytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
            stream.Write(boundaryBytes, 0, boundaryBytes.Length);

            stream.Close();

            // 开始发起请求，并获取服务器返回的结果。
            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), encoding))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private void testToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                // Read file data
                FileStream fs = new FileStream(@"E:\ThreeKingdoms\doc\静态数据表\buffer.xlsx", FileMode.Open, FileAccess.Read);
                byte[] data = new byte[fs.Length];
                fs.Read(data, 0, data.Length);
                fs.Close();

                //var Params = new Dictionary<string, string> { { "userid", "9" } };
                //UploadFilesToServer(new Uri("http://192.168.0.225:8000/albatross/res/SWFUpload/upload.php"), Params, "buffer.xlsx", "application/octet-stream", data);
                //UploadMultipart(data, "buffer.xlsx", "multipart/form-data", "http://192.168.0.225:8000/albatross/res/SWFUpload/upload.php");
                /*byte[] ret = cl.UploadFile("http://192.168.0.225:8000/albatross/res/SWFUpload/upload.php", @"E:\ThreeKingdoms\doc\静态数据表\buffer.xlsx");
                MessageBox.Show(ret.ToString());

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"http://192.168.0.225:8000/albatross/res/SWFUpload/upload.php");
                HttpWebResponse response = null;
                request.Method = "POST";
                request.ContentType = "multipart/form-data";
                request.AllowAutoRedirect = true;
                request.KeepAlive = true;
                response = (HttpWebResponse)request.GetResponse();
                StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                string content = sr.ReadToEnd();
                sr.Close();
                response.Close();*/

                string url = "http://192.168.0.225:8000/albatross/res/SWFUpload/upload.php";

                //only for extend key params
                Dictionary<string, string> keyvalues = new Dictionary<string, string>();
                keyvalues.Add("kkkk", "vvvv");

                Dictionary<string, string> fileList = new Dictionary<string, string>();
                //Filedata is important for server to recieve files
                fileList.Add("Filedata", @"E:\ThreeKingdoms\doc\静态数据表\buffer.xlsx");

                string resp = SendHttpRequestPost(url, keyvalues, fileList, Encoding.UTF8);
                MessageBox.Show(resp);

               /* // Generate post objects
                Dictionary<string, object> postParameters = new Dictionary<string, object>();
                postParameters.Add("filename", "People.doc");
                postParameters.Add("fileformat", "doc");
                postParameters.Add("file", new FormUpload.FileParameter(data, "People.doc", "application/msword"));

                // Create request and receive response
                string postURL = "http://192.168.0.225:8000/albatross/res/SWFUpload/upload.php";
                string userAgent = "Someone";
                HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(postURL, userAgent, postParameters);

                // Process response
                StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
                string fullResponse = responseReader.ReadToEnd();
                webResponse.Close();
                MessageBox.Show(fullResponse);*/

            }
            catch (Exception ex)
            {
                MessageBox.Show("Upload failed! " + ex.ToString());
            }

        }

        private void compareRaw(string format, DataGridView viewer, DataTable srcdt, DataTable destdt, Color markcolor)
        {
            if (format == "按单元格")
            {
                compareFiles(viewer, srcdt, destdt, markcolor);
            }
            else if (format == "按行")
            {
                compareFilesByRow(viewer, srcdt, destdt, markcolor);
            }
            else if (format == "按列")
            {
                compareFilesByCol(viewer, srcdt, destdt, markcolor);
            }
        }

        private void destdgview_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {

        }

        private void srcdgview_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {

        }

        //private void destdgview_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        //{
        //    e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        //}

        private void destdgview_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers  
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }

        private void srcdgview_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers  
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }

    }

    public static class FormUpload
    {
        private static readonly Encoding encoding = Encoding.UTF8;
        public static HttpWebResponse MultipartFormDataPost(string postUrl, string userAgent, Dictionary<string, object> postParameters)
        {
            string formDataBoundary = String.Format("----------{0:N}", Guid.NewGuid());
            string contentType = "multipart/form-data; boundary=" + formDataBoundary;

            byte[] formData = GetMultipartFormData(postParameters, formDataBoundary);

            return PostForm(postUrl, userAgent, contentType, formData);
        }
        private static HttpWebResponse PostForm(string postUrl, string userAgent, string contentType, byte[] formData)
        {
            HttpWebRequest request = WebRequest.Create(postUrl) as HttpWebRequest;

            if (request == null)
            {
                throw new NullReferenceException("request is not a http request");
            }

            // Set up the request properties.
            request.Method = "POST";
            request.ContentType = contentType;
            request.UserAgent = userAgent;
            request.CookieContainer = new CookieContainer();
            request.ContentLength = formData.Length;

            // You could add authentication here as well if needed:
            // request.PreAuthenticate = true;
            // request.AuthenticationLevel = System.Net.Security.AuthenticationLevel.MutualAuthRequested;
            // request.Headers.Add("Authorization", "Basic " + Convert.ToBase64String(System.Text.Encoding.Default.GetBytes("username" + ":" + "password")));

            // Send the form data to the request.
            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(formData, 0, formData.Length);
                requestStream.Close();
            }
            WebResponse wr = request.GetResponse();


            return request.GetResponse() as HttpWebResponse;
        }

        private static byte[] GetMultipartFormData(Dictionary<string, object> postParameters, string boundary)
        {
            Stream formDataStream = new System.IO.MemoryStream();
            bool needsCLRF = false;

            foreach (var param in postParameters)
            {
                // Thanks to feedback from commenters, add a CRLF to allow multiple parameters to be added.
                // Skip it on the first parameter, add it to subsequent parameters.
                if (needsCLRF)
                    formDataStream.Write(encoding.GetBytes("\r\n"), 0, encoding.GetByteCount("\r\n"));

                needsCLRF = true;

                if (param.Value is FileParameter)
                {
                    FileParameter fileToUpload = (FileParameter)param.Value;

                    // Add just the first part of this param, since we will write the file data directly to the Stream
                    string header = string.Format("--{0}\r\nContent-Disposition: form-data; name=\"{1}\"; filename=\"{2}\"\r\nContent-Type: {3}\r\n\r\n",
                        boundary,
                        param.Key,
                        fileToUpload.FileName ?? param.Key,
                        fileToUpload.ContentType ?? "application/octet-stream");

                    formDataStream.Write(encoding.GetBytes(header), 0, encoding.GetByteCount(header));

                    // Write the file data directly to the Stream, rather than serializing it to a string.
                    formDataStream.Write(fileToUpload.File, 0, fileToUpload.File.Length);
                }
                else
                {
                    string postData = string.Format("--{0}\r\nContent-Disposition: form-data; name=\"{1}\"\r\n\r\n{2}",
                        boundary,
                        param.Key,
                        param.Value);
                    formDataStream.Write(encoding.GetBytes(postData), 0, encoding.GetByteCount(postData));
                }
            }

            // Add the end of the request.  Start with a newline
            string footer = "\r\n--" + boundary + "--\r\n";
            formDataStream.Write(encoding.GetBytes(footer), 0, encoding.GetByteCount(footer));

            // Dump the Stream into a byte[]
            formDataStream.Position = 0;
            byte[] formData = new byte[formDataStream.Length];
            formDataStream.Read(formData, 0, formData.Length);
            formDataStream.Close();

            return formData;
        }

        public class FileParameter
        {
            public byte[] File { get; set; }
            public string FileName { get; set; }
            public string ContentType { get; set; }
            public FileParameter(byte[] file) : this(file, null) { }
            public FileParameter(byte[] file, string filename) : this(file, filename, null) { }
            public FileParameter(byte[] file, string filename, string contenttype)
            {
                File = file;
                FileName = filename;
                ContentType = contenttype;
            }
        }
    }


}
