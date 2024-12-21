//This file is part of ExamPaper Factory
using System;
using System.Linq;
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

        private void ManualForm_SizeChanged(object sender, EventArgs e)
        {
            manualRichTextBox.Width = Width - 18;
            manualRichTextBox.Height = Height - 67;
        }

        private void ZoomMinus2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentZoomFactor = manualRichTextBox.ZoomFactor;
            if (currentZoomFactor >= 1.1f) manualRichTextBox.ZoomFactor = currentZoomFactor - 0.1f;
        }

        private void ZoomPlus2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentZoomFactor = manualRichTextBox.ZoomFactor;
            if (currentZoomFactor < 1.5f) manualRichTextBox.ZoomFactor = currentZoomFactor + 0.1f;
        }

        private void SelectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            manualRichTextBox.SelectAll();
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (manualRichTextBox.SelectedText.Length > 0)
            {
                copyToolStripMenuItem.Enabled = true;
                Clipboard.SetDataObject(manualRichTextBox.SelectedText);
            }
        }

        private void ManualRichTextBox_MouseCaptureChanged(object sender, EventArgs e)
        {
            if (manualRichTextBox.SelectedText.Length > 0) copyToolStripMenuItem.Enabled = true;
            else copyToolStripMenuItem.Enabled = false;
        }

        private void Quit2ToolStripMenuItem_Click(object sender, EventArgs e) { this.Close(); }

        private void ReSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentZoomFactor = manualRichTextBox.ZoomFactor;
            manualRichTextBox.ZoomFactor = 1.0f;
        }

        private void ReSize2RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentZoomFactor = manualRichTextBox.ZoomFactor;
            manualRichTextBox.ZoomFactor = 1.0f;
        }
    }
}
