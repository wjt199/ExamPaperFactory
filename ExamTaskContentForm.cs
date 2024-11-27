//This file is part of ExamPaper Factory
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    public partial class ExamTaskContentForm : Form
    {

        private readonly ExamPaperTask currentExamPaperTask = null;//用来接受需要显示的数据内容

        public ExamTaskContentForm(ExamPaperTask examPaperTask)
        {
            currentExamPaperTask = examPaperTask;
            InitializeComponent();
        }


        private void ExamTaskContentForm_Load(object sender, EventArgs e)
        {
            int questionSum = 0;//小题类型总数
            string groupNmae = "fixedQuestionGroupBox";
            for (int i = 0; i < currentExamPaperTask.FirstQuestionList.Count; i++)
            {
                GroupBox groupBox = new GroupBox();
                groupBox.Text = currentExamPaperTask.QuestionName[i] + " （每题" + currentExamPaperTask.QuestionGrade[i] + "分, 小题间空" + currentExamPaperTask.QuestionNextLineSpace[i] + "行）";
                groupBox.Name = groupNmae + (1 + i).ToString();
                groupBox.Location = new Point(3, 4 + 18 * i + 28 * questionSum);
                questionSum += currentExamPaperTask.FirstQuestionList[i].Count;//小题类型数求和
                groupBox.Width = 512;
                groupBox.Height = 12 + 28 * currentExamPaperTask.FirstQuestionList[i].Count;
                AddQuestionsContent(groupBox, currentExamPaperTask.FirstQuestionList[i], i + 1);
                this.Controls.Add(groupBox);
                this.ClientSize = new System.Drawing.Size(518, groupBox.Location.Y + groupBox.Height + 2);
            }
        }


        private void AddQuestionsContent(GroupBox groupBox, List<ExamPaperTask.QuestionElement> questionElements, int index)
        {
            for (int i = 0; i < questionElements.Count; i++)
            {
                //pathLabel
                Label path = new Label();
                path.AutoEllipsis = false;
                path.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                path.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
                path.Location = new Point(20, 17 + 28 * i);
                path.Size = new Size(436, 18);
                path.Name = "pathLabel" + index + "_" + i;
                path.Text = questionElements[i].path;
                path.AutoEllipsis = true;
                path.TextAlign = ContentAlignment.MiddleLeft;
                if (!Utilities.CheckQuestionStorageFormat(questionElements[i].path))
                {
                    path.ForeColor = Color.Red;//检查路径有问题，文字红色显示
                                               // Create the ToolTip and associate with the Form container.
                    ToolTip toolTip = new ToolTip();
                    // Set up the delays for the ToolTip.
                    toolTip.AutoPopDelay = 5000;
                    toolTip.InitialDelay = 500;
                    toolTip.ReshowDelay = 50;
                    // Force the ToolTip text to be displayed whether or not the form is active.
                    toolTip.ShowAlways = true;
                    // Set up the ToolTip text for the Button and Checkbox.
                    toolTip.SetToolTip(path, "题库有误，请调整后重试！");
                }

                //questionAmountLabel
                Label questionAmount = new Label();
                questionAmount.AutoEllipsis = false;
                questionAmount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                questionAmount.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
                questionAmount.Location = new Point(462, 17 + 28 * i);
                questionAmount.Size = new Size(30, 18);
                questionAmount.Name = "questionAmountLabel" + index + "_" + i;
                questionAmount.Text = questionElements[i].amount.ToString();
                if (Utilities.CheckQuestionStorageFormat(questionElements[i].path) && Utilities.CountQuestionNumFromQuestionStoragePath(questionElements[i].path) < questionElements[i].amount)
                {
                    questionAmount.ForeColor = Color.Red;
                    // Create the ToolTip and associate with the Form container.
                    ToolTip toolTip = new ToolTip();
                    // Set up the delays for the ToolTip.
                    toolTip.AutoPopDelay = 5000;
                    toolTip.InitialDelay = 500;
                    toolTip.ReshowDelay = 50;
                    // Force the ToolTip text to be displayed whether or not the form is active.
                    toolTip.ShowAlways = true;
                    string str = "题数:" + questionElements[i].amount.ToString() + "大于题库总数:" + Utilities.CountQuestionNumFromQuestionStoragePath(questionElements[i].path).ToString() + "，请调整后重试";
                    // Set up the ToolTip text for the Button and Checkbox.
                    toolTip.SetToolTip(questionAmount, str);
                }
                questionAmount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

                //fixedAmount_PerLabel
                Label fixedAmount_PerLabel = new Label();
                fixedAmount_PerLabel.AutoSize = true;
                fixedAmount_PerLabel.Location = new Point(493, 20 + 28 * i);
                fixedAmount_PerLabel.Size = new Size(59, 12);
                fixedAmount_PerLabel.Name = "fixedAmount_PerLabel" + index + "_" + i;
                fixedAmount_PerLabel.Text = "题";
                fixedAmount_PerLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;


                //一同添加进groupBox控件
                groupBox.Controls.Add(path);
                groupBox.Controls.Add(questionAmount);
                groupBox.Controls.Add(fixedAmount_PerLabel);
            }
        }

        /// <summary>
        /// 该界面退出时，清空应用信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExamTaskContentForm_FormClosing(object sender, FormClosingEventArgs e)
        { MainForm.etc = null; }
    }

}
