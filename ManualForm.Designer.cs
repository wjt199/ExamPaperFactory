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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ManualForm));
            this.manualRichTextBox = new System.Windows.Forms.RichTextBox();
            this.manualMenuStrip = new System.Windows.Forms.MenuStrip();
            this.quitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.zoomPlusToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.zoomMinusToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.manualMenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // manualRichTextBox
            // 
            this.manualRichTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.manualRichTextBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.manualRichTextBox.Location = new System.Drawing.Point(12, 28);
            this.manualRichTextBox.Name = "manualRichTextBox";
            this.manualRichTextBox.ShowSelectionMargin = true;
            this.manualRichTextBox.Size = new System.Drawing.Size(899, 441);
            this.manualRichTextBox.TabIndex = 0;
            this.manualRichTextBox.Text = resources.GetString("manualRichTextBox.Text");
            // 
            // manualMenuStrip
            // 
            this.manualMenuStrip.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.manualMenuStrip.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.manualMenuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.manualMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.quitToolStripMenuItem,
            this.zoomPlusToolStripMenuItem,
            this.zoomMinusToolStripMenuItem});
            this.manualMenuStrip.Location = new System.Drawing.Point(0, 0);
            this.manualMenuStrip.Name = "manualMenuStrip";
            this.manualMenuStrip.Size = new System.Drawing.Size(923, 25);
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
            // ManualForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(923, 481);
            this.Controls.Add(this.manualRichTextBox);
            this.Controls.Add(this.manualMenuStrip);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.manualMenuStrip;
            this.MaximizeBox = false;
            this.Name = "ManualForm";
            this.Text = "ExamPaper Factory 使用手册";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ManualForm_FormClosing);
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
    }
}