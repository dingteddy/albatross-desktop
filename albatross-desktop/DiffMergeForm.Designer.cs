namespace albatross_desktop
{
    partial class DiffMergeForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.destdgview = new System.Windows.Forms.DataGridView();
            this.srcdgview = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.比较ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.比较CToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.destdgview)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.srcdgview)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.destdgview, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.srcdgview, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 25);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1030, 602);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // destdgview
            // 
            this.destdgview.AllowDrop = true;
            this.destdgview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.destdgview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.destdgview.Location = new System.Drawing.Point(3, 3);
            this.destdgview.Name = "destdgview";
            this.destdgview.ReadOnly = true;
            this.destdgview.RowTemplate.Height = 23;
            this.destdgview.Size = new System.Drawing.Size(509, 596);
            this.destdgview.TabIndex = 0;
            this.destdgview.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.destdgview_CellEnter);
            this.destdgview.Scroll += new System.Windows.Forms.ScrollEventHandler(this.destdgview_Scroll);
            this.destdgview.DragDrop += new System.Windows.Forms.DragEventHandler(this.destdgview_DragDrop);
            this.destdgview.DragEnter += new System.Windows.Forms.DragEventHandler(this.destdgview_DragEnter);
            // 
            // srcdgview
            // 
            this.srcdgview.AllowDrop = true;
            this.srcdgview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.srcdgview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.srcdgview.Location = new System.Drawing.Point(518, 3);
            this.srcdgview.Name = "srcdgview";
            this.srcdgview.ReadOnly = true;
            this.srcdgview.RowTemplate.Height = 23;
            this.srcdgview.Size = new System.Drawing.Size(509, 596);
            this.srcdgview.TabIndex = 1;
            this.srcdgview.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.srcdgview_CellEnter);
            this.srcdgview.Scroll += new System.Windows.Forms.ScrollEventHandler(this.srcdgview_Scroll);
            this.srcdgview.DragDrop += new System.Windows.Forms.DragEventHandler(this.srcdgview_DragDrop);
            this.srcdgview.DragEnter += new System.Windows.Forms.DragEventHandler(this.srcdgview_DragEnter);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.比较ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1030, 25);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 比较ToolStripMenuItem
            // 
            this.比较ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.比较CToolStripMenuItem});
            this.比较ToolStripMenuItem.Name = "比较ToolStripMenuItem";
            this.比较ToolStripMenuItem.Size = new System.Drawing.Size(62, 21);
            this.比较ToolStripMenuItem.Text = "操作(&O)";
            // 
            // 比较CToolStripMenuItem
            // 
            this.比较CToolStripMenuItem.Name = "比较CToolStripMenuItem";
            this.比较CToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.R)));
            this.比较CToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.比较CToolStripMenuItem.Text = "比较(&R)";
            this.比较CToolStripMenuItem.Click += new System.EventHandler(this.比较CToolStripMenuItem_Click);
            // 
            // DiffMergeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1030, 627);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "DiffMergeForm";
            this.Text = "DiffMergeForm";
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.destdgview)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.srcdgview)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView destdgview;
        private System.Windows.Forms.DataGridView srcdgview;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 比较ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 比较CToolStripMenuItem;
    }
}