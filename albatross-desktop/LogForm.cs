using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace albatross_desktop
{
    public partial class LogForm : Form
    {
        StreamReader streamReader;
        public LogForm(StreamReader sr)
        {
            InitializeComponent();

            streamReader = sr;
            showlog();
        }

        private void showlog()
        {
            textBox1.Text = streamReader.ReadToEnd();
            textBox1.Focus();//获取焦点
            textBox1.Select(textBox1.TextLength, 0);//光标定位到文本最后
            textBox1.ScrollToCaret();//滚动到光标处
            streamReader.BaseStream.Seek(0, SeekOrigin.Begin);
        }

        private void textBox1_DoubleClick(object sender, EventArgs e)
        {
            showlog();
        }
    }
}
