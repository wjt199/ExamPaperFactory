namespace ExamPaperFactory
{
    partial class SpecialForm
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
            this.passwdTextBox = new System.Windows.Forms.TextBox();
            this.enterButton = new System.Windows.Forms.Button();
            this.quitButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // passwdTextBox
            // 
            this.passwdTextBox.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.passwdTextBox.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.passwdTextBox.Location = new System.Drawing.Point(12, 12);
            this.passwdTextBox.Name = "passwdTextBox";
            this.passwdTextBox.PasswordChar = ' ';
            this.passwdTextBox.Size = new System.Drawing.Size(174, 23);
            this.passwdTextBox.TabIndex = 1;
            this.passwdTextBox.WordWrap = false;
            // 
            // enterButton
            // 
            this.enterButton.Location = new System.Drawing.Point(12, 41);
            this.enterButton.Name = "enterButton";
            this.enterButton.Size = new System.Drawing.Size(75, 21);
            this.enterButton.TabIndex = 2;
            this.enterButton.Text = "确定";
            this.enterButton.UseVisualStyleBackColor = true;
            this.enterButton.Click += new System.EventHandler(this.EnterButton_Click);
            // 
            // quitButton
            // 
            this.quitButton.Location = new System.Drawing.Point(111, 41);
            this.quitButton.Name = "quitButton";
            this.quitButton.Size = new System.Drawing.Size(75, 21);
            this.quitButton.TabIndex = 3;
            this.quitButton.Text = "取消";
            this.quitButton.UseVisualStyleBackColor = true;
            this.quitButton.Click += new System.EventHandler(this.QuitButton_Click);
            // 
            // SpecialForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(200, 71);
            this.Controls.Add(this.quitButton);
            this.Controls.Add(this.enterButton);
            this.Controls.Add(this.passwdTextBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SpecialForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "调试进入口令";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox passwdTextBox;
        private System.Windows.Forms.Button enterButton;
        private System.Windows.Forms.Button quitButton;
    }
}