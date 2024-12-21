using System;
using System.Drawing;
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
            if (passwdTextBox.Text == "666") { IsPassedRight = true; this.Close(); }
            else
            {
                toolStripStatusLabel.ForeColor = Color.Red;
                toolStripStatusLabel.Text = "口令错误";
            }
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

        private void PasswdTextBox_TextChanged(object sender, EventArgs e)
        {
            if (passwdTextBox.Text == String.Empty)
            {
                toolStripStatusLabel.ForeColor = SystemColors.ControlText;
                toolStripStatusLabel.Text = "口令：666 (强进或异)";
            }
        }
    }
}
