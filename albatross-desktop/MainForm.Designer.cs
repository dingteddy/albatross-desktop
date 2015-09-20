namespace albatross_desktop
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.重新加载RToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.保存XMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.保存EXCELToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.编辑RToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.配置PToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.配置PToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.dgview = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.添加行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.添加列ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusInfo = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.转换工具VToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgview)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.编辑RToolStripMenuItem,
            this.配置PToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(785, 25);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem1,
            this.saveToolStripMenuItem,
            this.重新加载RToolStripMenuItem,
            this.保存XMLToolStripMenuItem,
            this.保存EXCELToolStripMenuItem});
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.openToolStripMenuItem.Text = "文件(&F)";
            // 
            // openToolStripMenuItem1
            // 
            this.openToolStripMenuItem1.Name = "openToolStripMenuItem1";
            this.openToolStripMenuItem1.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem1.Size = new System.Drawing.Size(204, 22);
            this.openToolStripMenuItem1.Text = "打开(&O)...";
            this.openToolStripMenuItem1.Click += new System.EventHandler(this.openBtn_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.saveToolStripMenuItem.Text = "另存为(&S)...";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // 重新加载RToolStripMenuItem
            // 
            this.重新加载RToolStripMenuItem.Name = "重新加载RToolStripMenuItem";
            this.重新加载RToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.重新加载RToolStripMenuItem.Text = "重新加载(&R)";
            // 
            // 保存XMLToolStripMenuItem
            // 
            this.保存XMLToolStripMenuItem.Name = "保存XMLToolStripMenuItem";
            this.保存XMLToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Alt) 
            | System.Windows.Forms.Keys.X)));
            this.保存XMLToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.保存XMLToolStripMenuItem.Text = "保存XML";
            this.保存XMLToolStripMenuItem.Click += new System.EventHandler(this.saveXMLToolStripMenuItem_Click);
            // 
            // 保存EXCELToolStripMenuItem
            // 
            this.保存EXCELToolStripMenuItem.Name = "保存EXCELToolStripMenuItem";
            this.保存EXCELToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Alt) 
            | System.Windows.Forms.Keys.E)));
            this.保存EXCELToolStripMenuItem.Size = new System.Drawing.Size(204, 22);
            this.保存EXCELToolStripMenuItem.Text = "保存EXCEL";
            this.保存EXCELToolStripMenuItem.Click += new System.EventHandler(this.saveEXCELToolStripMenuItem_Click);
            // 
            // 编辑RToolStripMenuItem
            // 
            this.编辑RToolStripMenuItem.Name = "编辑RToolStripMenuItem";
            this.编辑RToolStripMenuItem.Size = new System.Drawing.Size(60, 21);
            this.编辑RToolStripMenuItem.Text = "编辑(&R)";
            // 
            // 配置PToolStripMenuItem
            // 
            this.配置PToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.配置PToolStripMenuItem1,
            this.转换工具VToolStripMenuItem});
            this.配置PToolStripMenuItem.Name = "配置PToolStripMenuItem";
            this.配置PToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.配置PToolStripMenuItem.Text = "工具(&T)";
            // 
            // 配置PToolStripMenuItem1
            // 
            this.配置PToolStripMenuItem1.Name = "配置PToolStripMenuItem1";
            this.配置PToolStripMenuItem1.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P)));
            this.配置PToolStripMenuItem1.Size = new System.Drawing.Size(183, 22);
            this.配置PToolStripMenuItem1.Text = "配置(&P)";
            this.配置PToolStripMenuItem1.Click += new System.EventHandler(this.ConfigToolStripMenuItem1_Click);
            // 
            // dgview
            // 
            this.dgview.AllowUserToAddRows = false;
            this.dgview.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgview.ContextMenuStrip = this.contextMenuStrip1;
            this.dgview.Location = new System.Drawing.Point(153, 42);
            this.dgview.Name = "dgview";
            this.dgview.RowTemplate.Height = 23;
            this.dgview.Size = new System.Drawing.Size(48, 76);
            this.dgview.TabIndex = 4;
            this.dgview.Visible = false;
            this.dgview.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgview_CellBeginEdit);
            this.dgview.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgview_CellEndEdit);
            this.dgview.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgview_CellValueChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.添加行ToolStripMenuItem,
            this.添加列ToolStripMenuItem,
            this.删除行ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(130, 70);
            // 
            // 添加行ToolStripMenuItem
            // 
            this.添加行ToolStripMenuItem.Name = "添加行ToolStripMenuItem";
            this.添加行ToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.添加行ToolStripMenuItem.Text = "添加行(&A)";
            this.添加行ToolStripMenuItem.Click += new System.EventHandler(this.AddRowToolStripMenuItem_Click);
            // 
            // 添加列ToolStripMenuItem
            // 
            this.添加列ToolStripMenuItem.Name = "添加列ToolStripMenuItem";
            this.添加列ToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.添加列ToolStripMenuItem.Text = "添加列(&C)";
            // 
            // 删除行ToolStripMenuItem
            // 
            this.删除行ToolStripMenuItem.Name = "删除行ToolStripMenuItem";
            this.删除行ToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.删除行ToolStripMenuItem.Text = "删除行(&D)";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(299, 137);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(100, 23);
            this.progressBar1.TabIndex = 5;
            this.progressBar1.Visible = false;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusInfo,
            this.toolStripStatusLabel2,
            this.toolStripStatusLabel3});
            this.statusStrip1.Location = new System.Drawing.Point(0, 450);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(785, 22);
            this.statusStrip1.TabIndex = 6;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusInfo
            // 
            this.statusInfo.Name = "statusInfo";
            this.statusInfo.Size = new System.Drawing.Size(697, 17);
            this.statusInfo.Spring = true;
            this.statusInfo.Text = "info";
            this.statusInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.statusInfo.Click += new System.EventHandler(this.statusInfo_Click);
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(42, 17);
            this.toolStripStatusLabel2.Text = "status";
            // 
            // toolStripStatusLabel3
            // 
            this.toolStripStatusLabel3.Name = "toolStripStatusLabel3";
            this.toolStripStatusLabel3.Size = new System.Drawing.Size(31, 17);
            this.toolStripStatusLabel3.Text = "N/A";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Control;
            this.panel1.Controls.Add(this.dgview);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Location = new System.Drawing.Point(0, 267);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(333, 183);
            this.panel1.TabIndex = 7;
            this.panel1.Visible = false;
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(412, 102);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(250, 250);
            this.webBrowser1.TabIndex = 8;
            this.webBrowser1.Visible = false;
            // 
            // 转换工具VToolStripMenuItem
            // 
            this.转换工具VToolStripMenuItem.Name = "转换工具VToolStripMenuItem";
            this.转换工具VToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.T)));
            this.转换工具VToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.转换工具VToolStripMenuItem.Text = "转换工具(&T)";
            this.转换工具VToolStripMenuItem.Click += new System.EventHandler(this.convertVToolStripMenuItem_Click);
            // 
            // MainForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.ClientSize = new System.Drawing.Size(785, 472);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.MainForm_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.MainForm_DragEnter);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgview)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem1;
        private System.Windows.Forms.DataGridView dgview;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 添加行ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 添加列ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除行ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 配置PToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 配置PToolStripMenuItem1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusInfo;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel3;
        private System.Windows.Forms.ToolStripMenuItem 重新加载RToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 编辑RToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 保存XMLToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 保存EXCELToolStripMenuItem;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.ToolStripMenuItem 转换工具VToolStripMenuItem;

    }
}

