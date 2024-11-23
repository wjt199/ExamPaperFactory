//This file is part of ExamPaper Factory
using System;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    public partial class ManualForm : Form
    {
        public ManualForm()
        {
            InitializeComponent();
        }

        private void ZoomPlusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentZoomFactor = manualRichTextBox.ZoomFactor;
            if (currentZoomFactor < 1.5f) manualRichTextBox.ZoomFactor = currentZoomFactor + 0.1f;
        }

        private void ZoomMinusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentZoomFactor = manualRichTextBox.ZoomFactor;
            if (currentZoomFactor >= 1.1f) manualRichTextBox.ZoomFactor = currentZoomFactor - 0.1f;
            //Console.WriteLine(manualRichTextBox.ZoomFactor);
        }

        private void QuitToolStripMenuItem_Click(object sender, EventArgs e) { this.Close(); }

        private void ManualForm_FormClosing(object sender, FormClosingEventArgs e) { MainForm.mf = null; }
    }
}
