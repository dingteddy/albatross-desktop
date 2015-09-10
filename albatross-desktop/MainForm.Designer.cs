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
            this.dgview = new System.Windows.Forms.DataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.添加行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.添加列ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.配置PToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.配置PToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgview)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.配置PToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(853, 25);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem1,
            this.saveToolStripMenuItem});
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.openToolStripMenuItem.Text = "文件(&F)";
            // 
            // openToolStripMenuItem1
            // 
            this.openToolStripMenuItem1.Name = "openToolStripMenuItem1";
            this.openToolStripMenuItem1.Size = new System.Drawing.Size(127, 22);
            this.openToolStripMenuItem1.Text = "打开(&O)...";
            this.openToolStripMenuItem1.Click += new System.EventHandler(this.openBtn_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(127, 22);
            this.saveToolStripMenuItem.Text = "保存(&S)...";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // dgview
            // 
            this.dgview.AllowUserToAddRows = false;
            this.dgview.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgview.ContextMenuStrip = this.contextMenuStrip1;
            this.dgview.Location = new System.Drawing.Point(12, 392);
            this.dgview.Name = "dgview";
            this.dgview.RowTemplate.Height = 23;
            this.dgview.Size = new System.Drawing.Size(48, 76);
            this.dgview.TabIndex = 4;
            this.dgview.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(76, 419);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(100, 23);
            this.progressBar1.TabIndex = 5;
            this.progressBar1.Visible = false;
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
            // 配置PToolStripMenuItem
            // 
            this.配置PToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.配置PToolStripMenuItem1});
            this.配置PToolStripMenuItem.Name = "配置PToolStripMenuItem";
            this.配置PToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.配置PToolStripMenuItem.Text = "工具(&T)";
            // 
            // 配置PToolStripMenuItem1
            // 
            this.配置PToolStripMenuItem1.Name = "配置PToolStripMenuItem1";
            this.配置PToolStripMenuItem1.Size = new System.Drawing.Size(152, 22);
            this.配置PToolStripMenuItem1.Text = "配置(&P)";
            this.配置PToolStripMenuItem1.Click += new System.EventHandler(this.ConfigToolStripMenuItem1_Click);
            // 
            // MainForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.ClientSize = new System.Drawing.Size(853, 505);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgview);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.MainForm_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.MainForm_DragEnter);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgview)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
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

    }
}

