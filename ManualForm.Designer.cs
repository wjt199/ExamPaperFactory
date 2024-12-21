namespace ExamPaperFactory
{
    partial class ManualForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ManualForm));
            this.manualRichTextBox = new System.Windows.Forms.RichTextBox();
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.zoomPlus2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.zoomMinus2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.selectAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.quit2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.manualMenuStrip = new System.Windows.Forms.MenuStrip();
            this.quitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reSizeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.zoomPlusToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.zoomMinusToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reSize2RToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip.SuspendLayout();
            this.manualMenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // manualRichTextBox
            // 
            this.manualRichTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.manualRichTextBox.ContextMenuStrip = this.contextMenuStrip;
            this.manualRichTextBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.manualRichTextBox.Location = new System.Drawing.Point(1, 27);
            this.manualRichTextBox.Name = "manualRichTextBox";
            this.manualRichTextBox.ReadOnly = true;
            this.manualRichTextBox.ShowSelectionMargin = true;
            this.manualRichTextBox.Size = new System.Drawing.Size(922, 453);
            this.manualRichTextBox.TabIndex = 0;
            this.manualRichTextBox.Text = resources.GetString("manualRichTextBox.Text");
            this.manualRichTextBox.MouseCaptureChanged += new System.EventHandler(this.ManualRichTextBox_MouseCaptureChanged);
            // 
            // contextMenuStrip
            // 
            this.contextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.zoomPlus2ToolStripMenuItem,
            this.reSize2RToolStripMenuItem,
            this.zoomMinus2ToolStripMenuItem,
            this.selectAllToolStripMenuItem,
            this.copyToolStripMenuItem,
            this.quit2ToolStripMenuItem});
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new System.Drawing.Size(188, 158);
            // 
            // zoomPlus2ToolStripMenuItem
            // 
            this.zoomPlus2ToolStripMenuItem.Name = "zoomPlus2ToolStripMenuItem";
            this.zoomPlus2ToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Up)));
            this.zoomPlus2ToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.zoomPlus2ToolStripMenuItem.Text = "放大(U)";
            this.zoomPlus2ToolStripMenuItem.Click += new System.EventHandler(this.ZoomPlus2ToolStripMenuItem_Click);
            // 
            // zoomMinus2ToolStripMenuItem
            // 
            this.zoomMinus2ToolStripMenuItem.Name = "zoomMinus2ToolStripMenuItem";
            this.zoomMinus2ToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Down)));
            this.zoomMinus2ToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.zoomMinus2ToolStripMenuItem.Text = "缩小(D)";
            this.zoomMinus2ToolStripMenuItem.Click += new System.EventHandler(this.ZoomMinus2ToolStripMenuItem_Click);
            // 
            // selectAllToolStripMenuItem
            // 
            this.selectAllToolStripMenuItem.Name = "selectAllToolStripMenuItem";
            this.selectAllToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.A)));
            this.selectAllToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.selectAllToolStripMenuItem.Text = "全选(A)";
            this.selectAllToolStripMenuItem.Click += new System.EventHandler(this.SelectAllToolStripMenuItem_Click);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Enabled = false;
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.C)));
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.copyToolStripMenuItem.Text = "复制(C)";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.CopyToolStripMenuItem_Click);
            // 
            // quit2ToolStripMenuItem
            // 
            this.quit2ToolStripMenuItem.Name = "quit2ToolStripMenuItem";
            this.quit2ToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.quit2ToolStripMenuItem.Text = "退出(Q)";
            this.quit2ToolStripMenuItem.Click += new System.EventHandler(this.Quit2ToolStripMenuItem_Click);
            // 
            // manualMenuStrip
            // 
            this.manualMenuStrip.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.manualMenuStrip.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.manualMenuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.manualMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.quitToolStripMenuItem,
            this.reSizeToolStripMenuItem,
            this.zoomPlusToolStripMenuItem,
            this.zoomMinusToolStripMenuItem});
            this.manualMenuStrip.Location = new System.Drawing.Point(0, 0);
            this.manualMenuStrip.Name = "manualMenuStrip";
            this.manualMenuStrip.Size = new System.Drawing.Size(924, 25);
            this.manualMenuStrip.TabIndex = 1;
            // 
            // quitToolStripMenuItem
            // 
            this.quitToolStripMenuItem.Name = "quitToolStripMenuItem";
            this.quitToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Q)));
            this.quitToolStripMenuItem.Size = new System.Drawing.Size(62, 21);
            this.quitToolStripMenuItem.Text = "退出(Q)";
            this.quitToolStripMenuItem.Click += new System.EventHandler(this.QuitToolStripMenuItem_Click);
            // 
            // reSizeToolStripMenuItem
            // 
            this.reSizeToolStripMenuItem.Name = "reSizeToolStripMenuItem";
            this.reSizeToolStripMenuItem.Size = new System.Drawing.Size(60, 21);
            this.reSizeToolStripMenuItem.Text = "复位(R)";
            this.reSizeToolStripMenuItem.Click += new System.EventHandler(this.ReSizeToolStripMenuItem_Click);
            // 
            // zoomPlusToolStripMenuItem
            // 
            this.zoomPlusToolStripMenuItem.Name = "zoomPlusToolStripMenuItem";
            this.zoomPlusToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Up)));
            this.zoomPlusToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.zoomPlusToolStripMenuItem.Text = "放大(↑)";
            this.zoomPlusToolStripMenuItem.Click += new System.EventHandler(this.ZoomPlusToolStripMenuItem_Click);
            // 
            // zoomMinusToolStripMenuItem
            // 
            this.zoomMinusToolStripMenuItem.Name = "zoomMinusToolStripMenuItem";
            this.zoomMinusToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.zoomMinusToolStripMenuItem.Text = "缩小(↓)";
            this.zoomMinusToolStripMenuItem.Click += new System.EventHandler(this.ZoomMinusToolStripMenuItem_Click);
            // 
            // reSize2RToolStripMenuItem
            // 
            this.reSize2RToolStripMenuItem.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);
            this.reSize2RToolStripMenuItem.Name = "reSize2RToolStripMenuItem";
            this.reSize2RToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.R)));
            this.reSize2RToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.reSize2RToolStripMenuItem.Text = "复位(R)";
            this.reSize2RToolStripMenuItem.Click += new System.EventHandler(this.ReSize2RToolStripMenuItem_Click);
            // 
            // ManualForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(924, 481);
            this.Controls.Add(this.manualRichTextBox);
            this.Controls.Add(this.manualMenuStrip);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.manualMenuStrip;
            this.MaximizeBox = false;
            this.Name = "ManualForm";
            this.Text = "ExamPaper Factory 使用手册";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ManualForm_FormClosing);
            this.SizeChanged += new System.EventHandler(this.ManualForm_SizeChanged);
            this.contextMenuStrip.ResumeLayout(false);
            this.manualMenuStrip.ResumeLayout(false);
            this.manualMenuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox manualRichTextBox;
        private System.Windows.Forms.MenuStrip manualMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem quitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem zoomPlusToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem zoomMinusToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem zoomPlus2ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem zoomMinus2ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem selectAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem quit2ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reSizeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reSize2RToolStripMenuItem;
    }
}
