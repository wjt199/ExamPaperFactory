using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    }
}
