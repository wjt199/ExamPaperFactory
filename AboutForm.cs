//This file is part of ExamPaper Factory
using System;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    public partial class AboutForm : Form
    {

        //FormStartPosition fatherStartPosition = FormStartPosition.Manual;

        public AboutForm()
        {
            InitializeComponent();
        }

        private void QuitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AboutForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            MainForm.af = null;
        }
    }
}
