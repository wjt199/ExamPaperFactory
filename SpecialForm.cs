using System;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    public partial class SpecialForm : Form
    {
        bool isPassedRight = false;
        public SpecialForm()
        {
            InitializeComponent();
        }
        public bool IsPassedRight { get => isPassedRight; set => isPassedRight = value; }

        private void EnterButton_Click(object sender, EventArgs e)
        {
            if (passwdTextBox.Text == "1234567890") IsPassedRight = true;
            this.Close();
        }

        private void QuitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 按Enter键直接确认
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">传递按键信息</param>
        private void PasswdTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    EnterButton_Click(sender, e);
                    break;
                default:
                    break;
            }
        }
    }
}
