namespace ExamPaperFactory
{
    partial class AboutForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutForm));
            this.nameTextLabel = new System.Windows.Forms.Label();
            this.copyRightTextLabel = new System.Windows.Forms.Label();
            this.authorLabel = new System.Windows.Forms.Label();
            this.quitButton = new System.Windows.Forms.Button();
            this.examPaperFactoryPictureBox = new System.Windows.Forms.PictureBox();
            this.telLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.examPaperFactoryPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // nameTextLabel
            // 
            this.nameTextLabel.AutoSize = true;
            this.nameTextLabel.Font = new System.Drawing.Font("Times New Roman", 9F);
            this.nameTextLabel.Location = new System.Drawing.Point(117, 32);
            this.nameTextLabel.Name = "nameTextLabel";
            this.nameTextLabel.Size = new System.Drawing.Size(228, 15);
            this.nameTextLabel.TabIndex = 0;
            this.nameTextLabel.Text = "ExamPaper Factory 1.0.4Alpha -[2024/11/29]";
            // 
            // copyRightTextLabel
            // 
            this.copyRightTextLabel.AutoSize = true;
            this.copyRightTextLabel.Font = new System.Drawing.Font("Times New Roman", 9F);
            this.copyRightTextLabel.Location = new System.Drawing.Point(117, 49);
            this.copyRightTextLabel.Name = "copyRightTextLabel";
            this.copyRightTextLabel.Size = new System.Drawing.Size(250, 15);
            this.copyRightTextLabel.TabIndex = 1;
            this.copyRightTextLabel.Text = "Copyright © 2024 Nineteen21. All rights reserved.";
            // 
            // authorLabel
            // 
            this.authorLabel.AutoSize = true;
            this.authorLabel.Font = new System.Drawing.Font("Times New Roman", 9F);
            this.authorLabel.Location = new System.Drawing.Point(117, 66);
            this.authorLabel.Name = "authorLabel";
            this.authorLabel.Size = new System.Drawing.Size(142, 15);
            this.authorLabel.TabIndex = 2;
            this.authorLabel.Text = "Developed by：Nineteen21";
            // 
            // quitButton
            // 
            this.quitButton.Location = new System.Drawing.Point(153, 111);
            this.quitButton.Name = "quitButton";
            this.quitButton.Size = new System.Drawing.Size(78, 23);
            this.quitButton.TabIndex = 3;
            this.quitButton.Text = "确定";
            this.quitButton.UseVisualStyleBackColor = true;
            this.quitButton.Click += new System.EventHandler(this.QuitButton_Click);
            // 
            // examPaperFactoryPictureBox
            // 
            this.examPaperFactoryPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("examPaperFactoryPictureBox.Image")));
            this.examPaperFactoryPictureBox.Location = new System.Drawing.Point(32, 32);
            this.examPaperFactoryPictureBox.Name = "examPaperFactoryPictureBox";
            this.examPaperFactoryPictureBox.Size = new System.Drawing.Size(64, 64);
            this.examPaperFactoryPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.examPaperFactoryPictureBox.TabIndex = 4;
            this.examPaperFactoryPictureBox.TabStop = false;
            // 
            // telLabel
            // 
            this.telLabel.AutoSize = true;
            this.telLabel.Font = new System.Drawing.Font("Times New Roman", 9F);
            this.telLabel.Location = new System.Drawing.Point(117, 81);
            this.telLabel.Name = "telLabel";
            this.telLabel.Size = new System.Drawing.Size(131, 15);
            this.telLabel.TabIndex = 5;
            this.telLabel.Text = "E-mail: wjt199@sina.com";
            // 
            // AboutForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(372, 154);
            this.Controls.Add(this.telLabel);
            this.Controls.Add(this.examPaperFactoryPictureBox);
            this.Controls.Add(this.quitButton);
            this.Controls.Add(this.authorLabel);
            this.Controls.Add(this.copyRightTextLabel);
            this.Controls.Add(this.nameTextLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 30);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "关于";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AboutForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.examPaperFactoryPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label nameTextLabel;
        private System.Windows.Forms.Label copyRightTextLabel;
        private System.Windows.Forms.Label authorLabel;
        private System.Windows.Forms.Button quitButton;
        private System.Windows.Forms.PictureBox examPaperFactoryPictureBox;
        private System.Windows.Forms.Label telLabel;
    }
}