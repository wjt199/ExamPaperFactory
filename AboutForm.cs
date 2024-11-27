//This file is part of ExamPaper Factory
using System;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    public partial class AboutForm : Form
    {
        public AboutForm() { InitializeComponent(); }

        /// <summary>
        /// 确定也就是退出按钮作用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void QuitButton_Click(object sender, EventArgs e) { this.Close(); }

        /// <summary>
        /// 窗体退出后方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AboutForm_FormClosing(object sender, FormClosingEventArgs e) { MainForm.af = null; }
    }
}
