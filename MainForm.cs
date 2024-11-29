//This file is part of ExamPaper Factory
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    //【1】第一步委托定义
    public delegate void CurrentTaskDelegate(ExamPaperTask examPaperTask);

    public partial class MainForm : Form
    {
        private string examPaperTaskFilePath; //任务文件绝对路径
        private ExamPaperTask currentTask = null;//当前任务
        private string examPaperDirPath;//试卷生成位置目录

        private ExamPaper examPaper = null;
        private BackgroundWorker backgroundWorker = null;

        public static NewExamTaskContentForm netc = null;//新建任务窗体
        public static EditExamTaskContentForm eetc = null;//编辑任务窗体
        public static ExamTaskContentForm etc = null;//更多内容窗体
        public static AboutForm af = null;//关于内容窗体
        public static ManualForm mf = null;//使用手册窗体窗体

        public bool isCurrentTaskExist = false;//判断当前界面是否有任务，右任务才涉及到任务新旧的问题
        public bool isCurrentTaskNew = false;//是新的任务文件标志（只有编辑过未被保存的文件才算新任务文件）
        public bool isCurrentRight;//当前任务文件经过检查后，修改此处的值，若处理后发现有问题，则这里为false；处理后没有问题则为true；

        public WorkResult workResult = null;

        public string ExamPaperTaskFilePath { get => examPaperTaskFilePath; set => examPaperTaskFilePath = value; }
        public ExamPaperTask CurrentTask { get => currentTask; set => currentTask = value; }
        public string ExamPaperFilePath { get => examPaperDirPath; set => examPaperDirPath = value; }
        internal ExamPaper ExamPaper { get => examPaper; set => examPaper = value; }

        public MainForm() { InitializeComponent(); }

        /// <summary>
        /// 窗体初始化（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            if (CurrentTask != null) this.moreExamPaperInfoButton.Enabled = true;
            if (File.Exists(@"configPath.txt"))
            {
                string[] config = File.ReadAllLines(@"configPath.txt", Encoding.UTF8);
                bool exampaper = false;
                for (int i = 0; i < config.Length; i++)
                {
                    if (config[i].Contains("exampaper"))
                    {
                        exampaper = true;
                        examPaperDirToolStripTextBox.Text = config[i].Substring(11);
                        break;
                    }
                }

                if (Directory.Exists(examPaperDirToolStripTextBox.Text))
                {
                    examPaperDirPath = examPaperDirToolStripTextBox.Text;
                    this.openExamPaperDirToolStripMenuItem.Enabled = true;
                }

                if (!exampaper) examPaperDirToolStripTextBox.Text = "点击↑↑↑选择试卷生成目录";
            }
            this.backgroundWorker = new BackgroundWorker();
        }

        /// <summary>
        /// 1.打开试卷任务文件，并将最近一次打开的文件位置存储到ConfigPath.txt文件中,
        /// 下次再次打开时优先获取ConfigPath.txt文件中相应字段的路径文件
        /// 2.将选中的xml文件解析出来，存放在内存中，并将内存中的内容显示在主界面上
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenEPFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isCurrentTaskNew) //如果还有文件没有保存是新文件
            {
                //提醒丢弃或取消
                DialogResult dialogResult = MessageBox.Show("是否丢弃当前任务并打开新任务？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Cancel) return;
            }
            string examPaperTaskFileDir; //任务文件所在目录绝对路径
            string examPaperTaskFileName; //任务文件名称
            //string type = "ExamPaperTaskFileDir";//目前不知道是干嘛用的为什么定义这个变量，后面需要的时候再说

            OpenFileDialog ofd = new OpenFileDialog();//打开文件
            ofd.Filter = "(*.xml)|*.xml";//设置过滤器，仅能选择xml文件类型
            //ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //先默认获取“我的电脑”路径，防止后面因为其他原因未能获取
            ofd.Title = "请选择试卷任务文件";//文件选择对话框标题
            ofd.Multiselect = false; //不允许多重选择
            ofd.RestoreDirectory = true;//记住上次打开的位置，即使软件重启也能记住，可完全替代下面的原始代码

            //选择目标文件后，提取出目标文件的绝对路径，同时将ConfigPath.txt缓存文件的内容更新为最新选择的文件所在目录位置
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ExamPaperTaskFilePath = ofd.FileName;//获取文件绝对路径
                examPaperTaskFileDir = Path.GetDirectoryName(ExamPaperTaskFilePath);//获取文件所在文件目录路径
                examPaperTaskFileName = Path.GetFileName(ExamPaperTaskFilePath);//获取文件名
                //Console.WriteLine("文件路径：{0}，文件名：{1}，文件夹：{2}", ExamPaperTaskFilePath, examPaperTaskFileName, examPaperTaskFileDir);
                //ConfigPath.updateConfigPath(type, examPaperTaskFileDir);//更新缓存信息，以便下次打开方便
            }
            else return;

            //处理xml文档之读取xml文档内容至对象examPaperTast中
            CurrentTask = new ExamPaperTask();
            if (CurrentTask.ReadXMLFile(ExamPaperTaskFilePath))//将读取到的数据投射到主界面上
            {
                DisplayInfoToMainForm(CurrentTask);
                isCurrentTaskNew = false;
                UpdateMainFormTiltleName();
            }
            else { DialogResult Error = System.Windows.Forms.MessageBox.Show("文件加载错误，请检查后重试", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            //检查更新主界面名称
        }

        /// <summary>
        /// 新建一个文件，并创建一个新的窗体用来输入文件相关的参数信息，同时需要使用委托传回所有合法的参数信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewEPFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isCurrentTaskNew) //如果还有文件没有保存是新文件
            {
                //提醒丢弃或取消
                DialogResult dialogResult = MessageBox.Show("是否丢弃当前任务并打开新任务？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Cancel) return;
            }

            if (netc == null)
            {
                netc = new NewExamTaskContentForm();
                netc.StartPosition = FormStartPosition.Manual;
                netc.Location = new System.Drawing.Point(this.Location.X, this.Location.Y + 55);
                netc.currentTaskDelegate = this.NewTaskFileReceiverAndDisplay; //【4】将第二个窗体中的委托变量和方法关联
                netc.Show();
            }
            else netc.Activate();

        }

        /// <summary>
        /// 编辑当前试卷任务文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void EditEPFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (eetc == null)
            {
                eetc = new EditExamTaskContentForm(CurrentTask);//把当前界面的任务传过去编辑，编辑完成后将数据返回来
                eetc.StartPosition = FormStartPosition.Manual;
                eetc.Location = new System.Drawing.Point(this.Location.X, this.Location.Y + 55);
                eetc.currentTaskDelegate = this.EditTaskFileReceiverAndDisplay; //【4】将第二个窗体中的委托变量和方法关联
                eetc.Show();
            }
            else eetc.Activate();
        }

        /// <summary>
        /// 退出按钮退出主程序（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void QuitToolStripMenuItem_Click(object sender, EventArgs e) { System.Environment.Exit(0); }

        /// <summary>
        /// 检查并更新主界面标题名称，规则如下：
        /// 0.程序刚加载时，显示：软件名称
        /// 1.打开试卷任务文件时，显示：软件名称 + 试卷问卷绝对路径；
        /// 2.新建的文件则显示：软件名称 + “*”
        /// 3.打开的试卷编辑完成后，显示：软件名称 + 试卷问卷绝对路径 + “*”
        /// 4.新建的文件在没有保存的情况下，编辑显示：软件名称 + “*”
        /// 5.新建的文件保存后则，则显示：软件名称 + 试卷问卷绝对路径（相当于打开的试卷任务文件）
        /// </summary>
        private void UpdateMainFormTiltleName()
        {
            string mainFormTitleHere = null;
            if (examPaperTaskFilePath != null) mainFormTitleHere = Path.GetFileName(examPaperTaskFilePath);
            if (!isCurrentTaskNew && examPaperTaskFilePath == null) this.Text = "试卷工厂";
            else if (!isCurrentTaskNew && examPaperTaskFilePath != null) this.Text = "试卷工厂(" + mainFormTitleHere + ")";
            else if (isCurrentTaskNew && examPaperTaskFilePath == null) this.Text = "试卷工厂*";
            else if (isCurrentTaskNew && examPaperTaskFilePath != null) this.Text = "试卷工厂(*" + mainFormTitleHere + ")";
        }

        /// <summary>
        /// 将ExamPaperTask对象值现实到界面上，并计算总分，同时视情激活更多按钮按键（已更新）
        /// </summary>
        /// <param name="examPaperTask"></param>
        private void DisplayInfoToMainForm(ExamPaperTask examPaperTask)
        {
            this.titleLable.Text = examPaperTask.TitleName;//试卷标题赋值

            this.paperSizeLabel.Text = examPaperTask.PaperSize;//试卷大小赋值
            this.textColumnsLabel.Text = examPaperTask.TextColumns;//试卷分栏
            this.orientationLabel.Text = examPaperTask.Orientation;//试卷排版赋值

            //试卷上下左边边距赋值
            this.marginTopLabel.Text = examPaperTask.MarginTop.ToString();
            this.marginBottomLabel.Text = examPaperTask.MarginBottom.ToString();
            this.marginLeftLabel.Text = examPaperTask.MarginLeft.ToString();
            this.marginRightLabel.Text = examPaperTask.MarginRight.ToString();

            //试卷标题格式
            this.mainTitleFontLabel.Text = examPaperTask.MainTitleFont;
            this.mainTitleFontStyleLabel.Text = examPaperTask.MainTitleFontStyle;
            this.mainTitleFontSizeLabel.Text = examPaperTask.MainTitleFontSize;
            this.mainTitleNextSpacingLineLabel.Text = examPaperTask.MainTitleNextLineSpace;

            //次标题格式单位姓名
            this.subTitleFontLabel.Text = examPaperTask.SubTitleFont;
            this.subTitleFontStyleLabel.Text = examPaperTask.SubTitleFontStyle;
            this.subTitleFontSizeLabel.Text = examPaperTask.SubTitleFontSize;
            this.subTitleNextSpacingLineLabel.Text = examPaperTask.SubTitleNextLineSpace;

            //一级标题格式
            this.firstHeadingFontLabel.Text = examPaperTask.FirstHeadingTitleFont;
            this.firstHeadingFontStyleLabel.Text = examPaperTask.FirstHeadingTitleFontStyle;
            this.firstHeadingFontSizeLabel.Text = examPaperTask.FirstHeadingTitleFontSize;
            this.firstHeadingNextSpacingLineLabel.Text = examPaperTask.FirstHeadingTitleNextLineSpace;

            //二级标题格式
            this.secondHeadingFontLabel.Text = examPaperTask.SecondHeadingTitleFont;
            this.secondHeadingFontStyleLabel.Text = examPaperTask.SecondHeadingTitleFontStyle;
            this.secondHeadingFontSizeLabel.Text = examPaperTask.SecondHeadingTitleFontSize;
            this.secondHeadingNextSpacingLineLabel.Text = examPaperTask.SecondHeadingTitleNextLineSpace;

            //大题数
            this.firstHeadingAmountLabel.Text = examPaperTask.FirstQuestionList.Count.ToString();

            //总分
            List<float> list = new List<float>();
            //Console.WriteLine("examPaperTask.FirstQuestionList.Count = {0}", examPaperTask.FirstQuestionList.Count);
            for (int i = 0; i < examPaperTask.FirstQuestionList.Count; i++)
            { foreach (ExamPaperTask.QuestionElement item in examPaperTask.FirstQuestionList[i]) list.Add(examPaperTask.QuestionGrade[i] * item.amount); }
            this.gradeAmountLabel.Text = list.Sum().ToString();

            //激活更多信息按钮
            if (examPaperTask.FirstQuestionList.Count != 0)
            {
                this.moreExamPaperInfoButton.Enabled = true;//有题目信息，更多信息按钮可用
                if (CheckQuestionListRight(examPaperTask)) //如果检查没有问题
                {
                    //右上角标志变绿
                    this.formExamPaperToolStripStatusLabel.ForeColor = Color.Green;
                    this.formExamPaperToolStripStatusLabel.Text = " √   数据正确";
                    this.formExamPaperToolStripStatusLabel.BorderStyle = Border3DStyle.Adjust;
                    this.formExamPaperToolStripStatusLabel.Font = new System.Drawing.Font("宋体", 9);
                    this.formExamPaperToolStripStatusLabel.Alignment = ToolStripItemAlignment.Left;
                    this.isCurrentRight = true;
                    if (Directory.Exists(this.examPaperDirToolStripTextBox.Text))
                    {
                        this.formExamPaperToolStripMenuItem.Enabled = true;
                        this.formExamPaperAndAnswerToolStripMenuItem.Enabled = true;
                    }

                    this.moreExamPaperInfoButton.ForeColor = SystemColors.ControlText;
                }
                else
                {
                    //右上角标志变红
                    this.isCurrentRight = false;
                    this.formExamPaperToolStripMenuItem.Enabled = false;
                    this.formExamPaperAndAnswerToolStripMenuItem.Enabled = false;
                    this.moreExamPaperInfoButton.ForeColor = Color.Red;
                    this.formExamPaperToolStripStatusLabel.ForeColor = Color.Red;
                    this.formExamPaperToolStripStatusLabel.Text = " ×   数据有误";
                    this.formExamPaperToolStripStatusLabel.BorderStyle = Border3DStyle.Adjust;
                    this.formExamPaperToolStripStatusLabel.Font = new System.Drawing.Font("宋体", 9);
                    this.formExamPaperToolStripStatusLabel.Alignment = ToolStripItemAlignment.Left;

                }
            }

            //检查如果当前任务有效，则激活“编辑任务”和“另存为”按钮
            if (CurrentTask != null)
            {
                this.editEPFileToolStripMenuItem.Enabled = true;//编辑按钮可用
                this.isCurrentTaskExist = true;
                this.saveAsToolStripMenuItem.Enabled = true;//另存为按钮可用
            }
        }

        /// <summary>
        /// 检查任务里面的题库路径及其对应的题目数量是否正确（不变）
        /// </summary>
        /// <param name="examPaperTask">任务文件</param>
        /// <returns>数量正确返回true；数量错误返回false</returns>
        private bool CheckQuestionListRight(ExamPaperTask examPaperTask)
        {

            for (int i = 0; i < examPaperTask.FirstQuestionList.Count; i++)
            {
                for (int j = 0; j < examPaperTask.FirstQuestionList[i].Count; j++)
                {
                    if (!File.Exists(examPaperTask.FirstQuestionList[i][j].path)) return false;
                    if (Utilities.CountQuestionNumFromQuestionStoragePath(examPaperTask.FirstQuestionList[i][j].path) < examPaperTask.FirstQuestionList[i][j].amount) return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 点击展示更多信息的按钮（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MoreExamPaperInfoButton_Click(object sender, EventArgs e)
        {
            if (etc == null)
            {
                etc = new ExamTaskContentForm(currentTask);
                etc.StartPosition = FormStartPosition.Manual;
                etc.Location = new System.Drawing.Point(this.Location.X, this.Location.Y + 55);
                etc.Show();
            }
            else etc.Activate();
        }

        /// <summary>
        /// 【2】编写委托相对应的方法:编辑后的文件，无需将打开文件路径删除
        /// 编辑完成后的数据不管数据是否发生了变化，默认当作发生了变化
        /// </summary>
        /// <param name="examPaperTask"></param>
        public void EditTaskFileReceiverAndDisplay(ExamPaperTask examPaperTask)
        {
            CurrentTask = examPaperTask;
            //ExamPaperTaskFilePath = null;//清空任务文件绝对路径，有就是有，没有就是没有
            isCurrentTaskNew = true;
            this.saveToolStripMenuItem.Enabled = true;
            DisplayInfoToMainForm(CurrentTask);
            UpdateMainFormTiltleName();
            this.Activate();//激活主窗体
        }

        /// <summary>
        /// 【2】编写委托相对应的方法：新建的文件，需要将打开文件路径删除
        /// 新打开的文件，默认是新文件
        /// </summary>
        /// <param name="examPaperTask"></param>
        public void NewTaskFileReceiverAndDisplay(ExamPaperTask examPaperTask)
        {
            CurrentTask = examPaperTask;
            ExamPaperTaskFilePath = null;//清空任务文件绝对路径
            isCurrentTaskNew = true;
            this.saveToolStripMenuItem.Enabled = true;
            DisplayInfoToMainForm(CurrentTask);
            UpdateMainFormTiltleName();
            this.Activate();//激活主窗体
        }

        /// <summary>
        /// 保存按钮（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!isCurrentTaskNew) return;//如果是旧文件直接退出

            if (ExamPaperTaskFilePath == null)
                SaveAs();
            else
            {
                //string saveFilePath = ExamPaperTaskFilePath; //获得文件路径
                //如果文件写入成功，将当前文件的文件路径换成保存的文件位置，并更新主界面标题
                if (CurrentTask.WriteToXMLFile(ExamPaperTaskFilePath))
                {
                    isCurrentTaskNew = false;//保存后就不是新文件了
                    saveToolStripMenuItem.Enabled = false;//保存按钮不可用
                    //ExamPaperTaskFilePath = saveFilePath;//更新文件路径
                    UpdateMainFormTiltleName();//更新主界面标题}
                }
                else
                {
                    DialogResult dialogResult = MessageBox.Show("文件保存失败", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (dialogResult == DialogResult.OK) return;
                }
            }
        }

        /// <summary>
        /// 另存为，把当前界面显示的试卷任务文件另存为一个新的（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FilterIndex = 1;//设置默认文件类型显示顺序 
            sfd.RestoreDirectory = true;//保存对话框是否记忆上次打开的目录,即使软件退出后再打开，他也能记住
            sfd.Filter = "试卷任务文件（*.xml）|*.xml";//设置文件类型
            sfd.FileName = "未命名";//默认文件名
            sfd.DefaultExt = "xml";//设置默认的文件名
            sfd.Title = "试卷任务文件另存为";//对话框标题

            //点了另存为按钮进入 
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = sfd.FileName.ToString(); //获得文件路径
                //如果文件写入成功，将当前文件的文件路径换成保存的文件位置，并更新主界面标题
                if (CurrentTask.WriteToXMLFile(saveFilePath))
                {
                    isCurrentTaskNew = false;//保存后就不是新文件了
                    this.saveToolStripMenuItem.Enabled = false;//保存按钮不可用
                    ExamPaperTaskFilePath = saveFilePath;//更新文件路径
                    UpdateMainFormTiltleName();//更新主界面标题
                }
            }
        }

        /// <summary>
        /// 另存为功能重载
        /// </summary>
        private void SaveAs()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FilterIndex = 1;//设置默认文件类型显示顺序 
            sfd.RestoreDirectory = true;//保存对话框是否记忆上次打开的目录,即使软件退出后再打开，他也能记住
            sfd.Filter = "试卷任务文件（*.xml）|*.xml";//设置文件类型
            sfd.FileName = "未命名";//默认文件名
            sfd.DefaultExt = "xml";//设置默认的文件名
            sfd.Title = "保存新试卷任务文件";//对话框标题

            //点了另存为按钮进入 
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = sfd.FileName.ToString(); //获得文件路径
                //如果文件写入成功，将当前文件的文件路径换成保存的文件位置，并更新主界面标题
                if (CurrentTask.WriteToXMLFile(saveFilePath))
                {
                    isCurrentTaskNew = false;//保存后就不是新文件了
                    this.saveToolStripMenuItem.Enabled = false;//保存按钮不可用
                    ExamPaperTaskFilePath = saveFilePath;//更新文件路径
                    UpdateMainFormTiltleName();//更新主界面标题
                }
            }
        }

        /// <summary>
        /// 选择试卷生成位置目录（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LocationExamPaperToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.RootFolder = Environment.SpecialFolder.MyComputer;
            //fbd.RootFolder = Environment.SpecialFolder.Personal;//只显示个人文档
            fbd.ShowNewFolderButton = true;//允许新建文件夹
            fbd.Description = "选择试卷生成位置";
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
                examPaperDirPath = fbd.SelectedPath;
                examPaperDirToolStripTextBox.Text = examPaperDirPath;
                if (isCurrentRight)
                {
                    this.formExamPaperToolStripMenuItem.Enabled = true;
                    this.formExamPaperAndAnswerToolStripMenuItem.Enabled = true;
                }
                this.openExamPaperDirToolStripMenuItem.Enabled = true;
                if (!File.Exists(@"configPath.txt"))
                {
                    File.WriteAllText(@"configPath.txt", "exampapers:" + examPaperDirToolStripTextBox.Text, Encoding.UTF8);
                    //MessageBox.Show("试卷路径保存成功");
                    //TextBoxConsoleLines.PrintLinesInConsole(this, "√ 试卷路径OK！", true);
                }
                else
                {
                    string[] config = File.ReadAllLines(@"configPath.txt", Encoding.UTF8);//读取"configPath.txt"文件
                    List<string> new_config = new List<string>();//新建"new_config"字符串组
                    bool exampaper = false;

                    if (config.Length != 0)//存在数据
                    {
                        foreach (string item in config)
                        {
                            if (!item.Contains("exampapers")) { new_config.Add(item); }//该行不包括关键字"exampapers"则复制该行
                            else
                            {
                                exampaper = true;
                                new_config.Add("exampapers:" + examPaperDirPath);//该行包括关键字"exampapers"则丢弃该行，并将新数据插入该行
                            }
                        }
                        if (!exampaper) { new_config.Add("exampapers:" + examPaperDirPath); }//搜索完都没有发现关键字"exampapers"，则直接添加该行
                    }
                    else { new_config.Add("exampapers:" + examPaperDirPath); }//不存在数据也是直接添加该行
                    File.WriteAllLines(@"configPath.txt", new_config.ToArray<string>());//将数据写入文件"configPath.txt"
                    //TextBoxConsoleLines.PrintLinesInConsole(this, "√ 试卷路径OK！", true);
                }
                //Form_epags.exampaperdir = true;
            }
        }

        /// <summary>
        /// 打开指定文件夹（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenExamPaperDirToolStripMenuItem_Click(object sender, EventArgs e)
        { if (examPaperDirPath != "") System.Diagnostics.Process.Start("Explorer.exe", examPaperDirPath); }


        /// <summary>
        /// 总分数值变化时进行判断，不为100时改变颜色进行提醒，不作为数据错误
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GradeAmountLabel_TextChanged(object sender, EventArgs e)
        {
            if (gradeAmountLabel.Text == "100") gradeAmountLabel.ForeColor = Color.Black;
            else gradeAmountLabel.ForeColor = Color.Gray;
        }

        #region 生成试卷和答案相关方法

        /// <summary>
        /// 生成试卷和答案按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormExamPaperAndAnswerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!backgroundWorker.IsBusy || backgroundWorker == null)
            {
                if (isCurrentRight)
                {
                    examPaper = new ExamPaper
                    {
                        ExamPaperTask1 = CurrentTask,//生成试卷所依据的任务文件为当前任务文件
                        ExamPaperDirPath = examPaperDirPath//给定试卷生成目录
                    };//创建试卷生成对象
                    examPaper.MainFormTitle = examPaperTaskFilePath != null ? examPaper.MainFormTitle = Path.GetFileNameWithoutExtension(examPaperTaskFilePath) : "";
                    examPaper.ExamPaperFileFormTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");//20241113210012(2024年11月13日21:00:12)
                    StartFormExamPaperAndAnswer();
                }
            }
        }

        /// <summary>
        /// 开始生成试卷和答案任务
        /// </summary>
        /// <param name="examPaper">试卷内容</param>
        private void StartFormExamPaperAndAnswer()
        {
            formExamPaperToolStripProgressBar.Visible = true;//使进度条可见
            formExamPaperToolStripStatusLabel.ForeColor = SystemColors.Desktop;//黑色
            //formExamPaperToolStripStatusLabel.Margin.Left = 140;
            backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,//可以报告进度
                WorkerSupportsCancellation = true//可以随时取消
            };//新建一个后台线程

            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_FormExamPaperAndAnswer);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_FormExamPaperAndAnswerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_FormExamPaperAndAnswer_ProgressChanged);
            backgroundWorker.RunWorkerAsync();
        }

        /// <summary>
        /// 进度条变更显示方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void BackgroundWorker_FormExamPaperAndAnswer_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            formExamPaperToolStripProgressBar.Value = e.ProgressPercentage;
            formExamPaperToolStripStatusLabel.Text = e.UserState as String;
        }

        /// <summary>
        /// 后台生成试卷和答案任务完成方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_FormExamPaperAndAnswerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //取得DoWork操作结果
            //WorkResult result = e.Result as WorkResult;
            if (e.Cancelled)
            {
                formExamPaperToolStripStatusLabel.ForeColor = Color.Red;//红色
                formExamPaperToolStripProgressBar.Visible = false;
                formExamPaperToolStripStatusLabel.Text = "错误：" + workResult.Result;
                formExamPaperToolStripStatusLabel.ToolTipText = "错误：" + workResult.Result;
                //Console.WriteLine("111");
                backgroundWorker.CancelAsync();
                //Console.WriteLine("222");
            }
            else
            {
                formExamPaperToolStripStatusLabel.ForeColor = Color.Green;//绿色
                formExamPaperToolStripProgressBar.Visible = false;
                formExamPaperToolStripStatusLabel.Text = "完成";
            }
            //两个变量在使用完后丢点
            workResult = null;
            backgroundWorker = new BackgroundWorker();
            examPaper = null;
        }

        /// <summary>
        /// 后台生成试卷和答案方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_FormExamPaperAndAnswer(object sender, DoWorkEventArgs e)
        {

            workResult = new WorkResult
            {
                Error = null,
                Result = null
            };//自定义返回结果类对象

            int progress = 0;

            progress = 1;
            backgroundWorker.ReportProgress(progress, "试卷生成中...");

            List<List<int[]>> chosenIndex = new List<List<int[]>>();//用来装试卷和答案序号的列表

            progress = 5;
            backgroundWorker.ReportProgress(progress, "试卷生成中...");

            //检查并创建生成试卷和答案的目录，反之非法行为导致生成失败
            if (!Directory.Exists(examPaperDirPath)) Directory.CreateDirectory(examPaperDirPath);

            ///生成试卷
            ExamPaper.ExamPaperName = ExamPaper.MainFormTitle + "（试卷）" + ExamPaper.ExamPaperFileFormTime + ".docx";
            //Console.WriteLine("ExamPaperName:{0}", ExamPaperName);

            //Word应用程序变量
            Microsoft.Office.Interop.Word.Application wordexampaperop = null;
            Document wordexampaper = null;

            progress = 8;
            backgroundWorker.ReportProgress(progress, "试卷生成中...");

            try
            {
                wordexampaperop = new Microsoft.Office.Interop.Word.Application();
                wordexampaperop.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;//屏蔽宏
                wordexampaperop.Visible = false;//使文档不可见,生成文件过程中文档不可见
                                                //wordexamop.Visible = true;//使文档可见

                Object exampapermissing = Missing.Value;//由于使用的是COM库，因此有许多变量需要使用Missing.Value代替

                wordexampaper = wordexampaperop.Documents.Add();

                //纸张大小设置
                if (ExamPaper.ExamPaperTask1.PaperSize == "A3") wordexampaper.PageSetup.PaperSize = WdPaperSize.wdPaperA3;//设置纸张样式为A3纸
                else if (ExamPaper.ExamPaperTask1.PaperSize == "A4") wordexampaper.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置纸张样式为A4纸
                else wordexampaper.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//默认纸张样式为A4规格

                //横竖版面设置
                if (ExamPaper.ExamPaperTask1.Orientation == "竖版") wordexampaper.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                else if (ExamPaper.ExamPaperTask1.Orientation == "横版") wordexampaper.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                else wordexampaper.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

                //页面设置
                wordexampaper.PageSetup.LeftMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginLeft.ToString()); //2.82CM
                wordexampaper.PageSetup.RightMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginRight.ToString());
                wordexampaper.PageSetup.TopMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginTop.ToString());//3.525CM
                wordexampaper.PageSetup.BottomMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginBottom.ToString());

                //分栏设置
                if (ExamPaper.ExamPaperTask1.TextColumns == "2") wordexampaper.PageSetup.TextColumns.SetCount(2);
                else wordexampaper.PageSetup.TextColumns.SetCount(1);

                //添加试卷标题
                wordexampaperop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordexampaperop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.MainTitleFontSize);
                wordexampaperop.Selection.Font.Name = ExamPaper.ExamPaperTask1.MainTitleFont;
                wordexampaperop.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;//首行缩进为0
                wordexampaperop.Selection.ParagraphFormat.FirstLineIndent = wordexampaper.Application.CentimetersToPoints((float)0.75 * 0);//首行缩进为0，全局设置为0
                wordexampaperop.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                wordexampaperop.Selection.ParagraphFormat.LineSpacing = float.Parse(ExamPaper.ExamPaperTask1.MainTitleNextLineSpace);
                wordexampaper.Paragraphs.Last.Range.Text = this.ExamPaper.ExamPaperTask1.TitleName + Environment.NewLine + Environment.NewLine;//试卷标题

                progress = 15;
                backgroundWorker.ReportProgress(progress, "试卷生成中...");

                object exampaperunite = WdUnits.wdStory;
                wordexampaperop.Selection.EndKey(ref exampaperunite, ref exampapermissing);//将光标移动到文本末尾
                wordexampaperop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.SubTitleFontSize);//次标题大小
                wordexampaperop.Selection.Font.Name = ExamPaper.ExamPaperTask1.SubTitleFont;
                wordexampaper.Paragraphs.Last.Range.Text = "单位：___________    姓名：___________    成绩：___________" + Environment.NewLine + Environment.NewLine;

                progress = 19;
                backgroundWorker.ReportProgress(progress, "试卷生成中...");

                int mount1 = ExamPaper.ExamPaperTask1.QuestionName.Count;
                int step1 = 30 / mount1;
                for (int i = 0; i < ExamPaper.ExamPaperTask1.QuestionName.Count; i++)
                {
                    //题目标题加内容
                    wordexampaperop.Selection.EndKey(ref exampaperunite, ref exampapermissing);//将光标移动到文本末尾
                    wordexampaperop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    wordexampaperop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.FirstHeadingTitleFontSize);//大题字体大小设置
                    wordexampaperop.Selection.Font.Name = ExamPaper.ExamPaperTask1.FirstHeadingTitleFont;//大题字体设置
                                                                                                         //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值,前面已经设置过了
                    wordexampaperop.Selection.ParagraphFormat.LineSpacing = float.Parse(ExamPaper.ExamPaperTask1.FirstHeadingTitleNextLineSpace);//大题行间距设置
                    float sum = 0;
                    for (int j = 0; j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count; j++) sum += ExamPaper.ExamPaperTask1.QuestionGrade[i] * ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].amount;
                    wordexampaper.Paragraphs.Last.Range.Text = Utilities.IntToBigNum(i + 1) + ExamPaper.ExamPaperTask1.QuestionName[i] + "（每题" + ExamPaper.ExamPaperTask1.QuestionGrade[i] + "分，共" + sum + "分）" + Environment.NewLine;
                    wordexampaperop.Selection.EndKey(ref exampaperunite, ref exampapermissing);//将光标移动到文本末尾
                    wordexampaperop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.SecondHeadingTitleFontSize);//小题字体大小设置
                    wordexampaperop.Selection.Font.Name = ExamPaper.ExamPaperTask1.SecondHeadingTitleFont;//小题字体设置
                                                                                                          //wordexamop.Selection.Font.;
                                                                                                          //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                    wordexampaperop.Selection.ParagraphFormat.LineSpacing = float.Parse(ExamPaper.ExamPaperTask1.SecondHeadingTitleNextLineSpace);//小题行间距设置

                    for (int j = 0; j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count; j++)
                    {
                        List<List<string>> questionList = new List<List<string>>();
                        //questionList = null;
                        List<int[]> ints = new List<int[]>();
                        while (j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count)
                        {
                            List<List<string>> newQuestionList = new List<List<string>>();
                            string questionStoragePath = ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].path;//题库路径
                            int num = ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].amount;//上述路径对应的选题数
                            int[] index = new int[num];
                            index = Utilities.GetNeededQuestionAndAnswerIndexFromQuestionStorage(questionStoragePath, num);
                            ints.Add(index);
                            newQuestionList = Utilities.GetNeededQuestionFromQuestionStorageByIndex(questionStoragePath, index);
                            questionList = Utilities.CombineList(questionList, newQuestionList);
                            j++;
                        }
                        chosenIndex.Add(ints);

                        for (int k = 0; k < questionList.Count; k++)
                        {
                            //Console.WriteLine("k:{0}", k);
                            for (int l = 0; l < questionList[k].Count; l++)
                            {
                                if (l == 0) wordexampaper.Paragraphs.Last.Range.Text = (k + 1).ToString() + ". " + questionList[k][l] + Environment.NewLine;
                                else wordexampaper.Paragraphs.Last.Range.Text = questionList[k][l] + Environment.NewLine;
                            }
                            for (int m = 0; m < ExamPaper.ExamPaperTask1.QuestionNextLineSpace[i]; m++)//每个大的题型的小题行距以每个大题第一个题库输入的行间距为准，后续题库的输入值不予采纳
                                wordexampaper.Paragraphs.Last.Range.Text = Environment.NewLine;
                        }

                    }
                    wordexampaper.Paragraphs.Last.Range.Text = Environment.NewLine;//大题之间的空行
                    progress = 19 + i * step1;
                    backgroundWorker.ReportProgress(progress, "试卷生成中...");
                }

                progress = 50;
                backgroundWorker.ReportProgress(progress, "试卷已完成");
                Thread.Sleep(1000);

                wordexampaper.SaveAs(Path.Combine(ExamPaper.ExamPaperDirPath, ExamPaper.ExamPaperName));//存储试卷
                wordexampaper.Close(ref exampapermissing, ref exampapermissing, ref exampapermissing);
                wordexampaper = null;
                wordexampaperop.Quit(ref exampapermissing, ref exampapermissing, ref exampapermissing);
                wordexampaperop = null;

                ///生成答案
                ExamPaper.ExamPaperAnswerName = ExamPaper.MainFormTitle + "（答案）" + ExamPaper.ExamPaperFileFormTime + ".docx";
                //Console.WriteLine("ExamPaperName:{0}", ExamPaperName);

                progress = 52;
                backgroundWorker.ReportProgress(progress, "答案生成中...");

                //Word应用程序变量
                Microsoft.Office.Interop.Word.Application wordexampaperanswerop = null;
                Document wordexampaperanser = null;

                progress = 55;
                backgroundWorker.ReportProgress(progress, "答案生成中...");

                wordexampaperanswerop = new Microsoft.Office.Interop.Word.Application();
                wordexampaperanswerop.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;//屏蔽宏
                wordexampaperanswerop.Visible = false;//使文档不可见,生成文件过程中文档不可见
                                                      //wordexamop.Visible = true;//使文档可见
                progress = 58;
                backgroundWorker.ReportProgress(progress, "答案生成中...");

                Object exampaperanswermissing = Missing.Value;//由于使用的是COM库，因此有许多变量需要使用Missing.Value代替

                wordexampaperanser = wordexampaperanswerop.Documents.Add();

                //纸张大小设置
                wordexampaperanser.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置纸张样式为A4纸
                //横竖版面设置
                wordexampaperanser.PageSetup.Orientation = WdOrientation.wdOrientPortrait;//设置为竖版
                //页面设置
                wordexampaperanser.PageSetup.LeftMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginLeft.ToString()); //2.82CM
                wordexampaperanser.PageSetup.RightMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginRight.ToString());
                wordexampaperanser.PageSetup.TopMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginTop.ToString());//3.525CM
                wordexampaperanser.PageSetup.BottomMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginBottom.ToString());

                progress = 60;
                backgroundWorker.ReportProgress(progress, "答案生成中...");

                //添加试卷标题
                wordexampaperanswerop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordexampaperanswerop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.MainTitleFontSize);
                wordexampaperanswerop.Selection.Font.Name = ExamPaper.ExamPaperTask1.MainTitleFont;
                wordexampaperanswerop.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;//首行缩进为0
                wordexampaperanswerop.Selection.ParagraphFormat.FirstLineIndent = wordexampaperanser.Application.CentimetersToPoints((float)0.75 * 0);//首行缩进为0，全局设置为0
                wordexampaperanswerop.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                wordexampaperanswerop.Selection.ParagraphFormat.LineSpacing = float.Parse(ExamPaper.ExamPaperTask1.MainTitleNextLineSpace);
                wordexampaperanser.Paragraphs.Last.Range.Text = this.ExamPaper.ExamPaperTask1.TitleName + "答案" + Environment.NewLine + Environment.NewLine;//试卷标题

                progress = 69;
                backgroundWorker.ReportProgress(progress, "答案生成中...");

                object exampaperanswerunite = WdUnits.wdStory;
                wordexampaperanswerop.Selection.EndKey(ref exampaperanswerunite, ref exampaperanswermissing);//将光标移动到文本末尾

                int step2 = 25 / mount1;

                for (int i = 0; i < ExamPaper.ExamPaperTask1.QuestionName.Count; i++)
                {
                    //题目标题加内容
                    wordexampaperanswerop.Selection.EndKey(ref exampaperanswerunite, ref exampaperanswermissing);
                    wordexampaperanswerop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    wordexampaperanswerop.Selection.Font.Size = 10;
                    wordexampaperanswerop.Selection.Font.Name = "黑体";
                    //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值,前面已经设置过了
                    wordexampaperanswerop.Selection.ParagraphFormat.LineSpacing = 14f;
                    wordexampaperanser.Paragraphs.Last.Range.Text = Utilities.IntToBigNum(i + 1) + ExamPaper.ExamPaperTask1.QuestionName[i] + Environment.NewLine;
                    wordexampaperanswerop.Selection.EndKey(ref exampaperanswerunite, ref exampaperanswermissing);//将光标移动到文本末尾
                    wordexampaperanswerop.Selection.Font.Size = 9;
                    wordexampaperanswerop.Selection.Font.Name = "宋体";//小答案内容字体设置
                    wordexampaperanswerop.Selection.ParagraphFormat.LineSpacing = 12f;//小题行间距设置
                    for (int j = 0; j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count; j++)
                    {
                        List<List<string>> answerList = new List<List<string>>();
                        answerList = null;
                        while (j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count)
                        {
                            List<List<string>> newAnswerList = new List<List<string>>();
                            string questionStoragePath = ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].path;
                            int num = ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].amount;
                            newAnswerList = Utilities.GetNeededAnswerFromQuestionStorageByIndex(questionStoragePath, chosenIndex[i][j]);
                            answerList = Utilities.CombineList(answerList, newAnswerList);
                            j++;
                        }

                        string answer = "";
                        for (int k = 0; k < answerList.Count; k++)
                        {
                            for (int l = 0; l < answerList[k].Count; l++)
                            {
                                if (l == 0) answer += "[" + (k + 1).ToString() + "]" + answerList[k][l] + "  ";
                                else answer += answerList[k][l] + "  ";
                            }
                        }
                        wordexampaperanser.Paragraphs.Last.Range.Text = answer + Environment.NewLine;


                    }

                    progress = 69 + i * step2;
                    backgroundWorker.ReportProgress(progress, "答案生成中...");
                }

                progress = 95;
                backgroundWorker.ReportProgress(progress, "答案已完成");
                Thread.Sleep(1000);

                wordexampaperanser.SaveAs(Path.Combine(ExamPaper.ExamPaperDirPath, ExamPaper.ExamPaperAnswerName));//存储试卷答案
                wordexampaperanser.Close(ref exampaperanswermissing, ref exampaperanswermissing, ref exampaperanswermissing);
                wordexampaperanser = null;
                wordexampaperanswerop.Quit(ref exampaperanswermissing, ref exampaperanswermissing, ref exampaperanswermissing);
                wordexampaperanswerop = null;

                progress = 99;
                Thread.Sleep(500);
                backgroundWorker.ReportProgress(progress, "正在进行最后步骤...");
                Thread.Sleep(500);
                progress = 100;
                backgroundWorker.ReportProgress(progress, "正在进行最后步骤...");
                Thread.Sleep(1000);
            }
            catch (Exception err)
            {
                //Console.WriteLine($"Processing failed: {err.Message}");
                if (progress < 50)
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "试卷生成错误");
                        Thread.Sleep(10);
                    }
                }
                else if (progress == 50)
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "试卷保存错误");
                        Thread.Sleep(10);
                    }
                }
                else if (progress < 95)
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "答案生成错误");
                        Thread.Sleep(10);
                    }
                }
                else if (progress == 95)
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "答案保存错误");
                        Thread.Sleep(10);
                    }
                }
                else
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "未可知其错");
                        Thread.Sleep(10);
                    }
                }
                if (err != null) { workResult.Error = err; workResult.Result = err.Message.Replace("\n", "").Replace("\r", "").Replace(Environment.NewLine, ""); e.Cancel = true; }
            }
        }

        #endregion

        #region 生成试卷方法 
        /// <summary>
        /// 生成试卷按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormExamPaperToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!backgroundWorker.IsBusy || backgroundWorker == null)
            {
                if (isCurrentRight)
                {
                    examPaper = new ExamPaper
                    {
                        ExamPaperTask1 = CurrentTask,//生成试卷所依据的任务文件为当前任务文件
                        ExamPaperDirPath = examPaperDirPath//给定试卷生成目录
                    };//创建试卷生成对象
                    examPaper.MainFormTitle = examPaperTaskFilePath != null ? examPaper.MainFormTitle = Path.GetFileNameWithoutExtension(examPaperTaskFilePath) : "";
                    examPaper.ExamPaperFileFormTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");//20241113210012(2024年11月13日21:00:12)

                    StartFormExamPaperWithoutAnswer();
                }
            }
        }

        /// <summary>
        /// 开始生成试卷
        /// </summary>
        private void StartFormExamPaperWithoutAnswer()
        {
            formExamPaperToolStripProgressBar.Visible = true;//使进度条可见
            formExamPaperToolStripStatusLabel.ForeColor = SystemColors.Desktop;//黑色
            backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,//可以报告进度
                WorkerSupportsCancellation = true//可以随时取消
            };//新建一个后台线程

            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_FormExamPaperWithoutAnswer);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_FormExamPaperWithoutAnswerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_FormExamPaperWithoutAnswer_ProgressChanged);
            backgroundWorker.RunWorkerAsync();
        }

        /// <summary>
        /// 进度条变更显示方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_FormExamPaperWithoutAnswer_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            formExamPaperToolStripProgressBar.Value = e.ProgressPercentage;
            formExamPaperToolStripStatusLabel.Text = e.UserState as String;
        }

        /// <summary>
        /// 后台生成试卷任务完成后方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_FormExamPaperWithoutAnswerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //if (e.Cancelled) formExamPaperToolStripStatusLabel.Text = "任务取消";

            //取得DoWork操作结果
            //WorkResult result = e.Result as WorkResult;
            if (e.Cancelled)
            {
                formExamPaperToolStripStatusLabel.ForeColor = Color.Red;//红色
                formExamPaperToolStripProgressBar.Visible = false;
                formExamPaperToolStripStatusLabel.Text = "错误：" + workResult.Result;
                formExamPaperToolStripStatusLabel.ToolTipText = "错误：" + workResult.Result;
                backgroundWorker.CancelAsync();
            }
            else
            {
                formExamPaperToolStripStatusLabel.ForeColor = Color.Green;//绿色
                formExamPaperToolStripProgressBar.Visible = false;
                formExamPaperToolStripStatusLabel.Text = "完成";
            }

            //使用完后丢掉
            backgroundWorker = new BackgroundWorker();
            examPaper = null;
            workResult = null;
        }

        /// <summary>
        /// 后台生成试卷方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_FormExamPaperWithoutAnswer(object sender, DoWorkEventArgs e)
        {
            workResult = new WorkResult
            {
                Error = null,
                Result = null
            };//自定义返回结果类对象

            int progress = 0;

            progress = 1;
            backgroundWorker.ReportProgress(progress, "试卷生成中...");

            //检查并创建生成试卷和答案的目录，反之非法行为导致生成失败
            if (!Directory.Exists(examPaperDirPath)) Directory.CreateDirectory(examPaperDirPath);
            ExamPaper.ExamPaperName = ExamPaper.MainFormTitle + "（试卷）" + ExamPaper.ExamPaperFileFormTime + ".docx";
            //Console.WriteLine("ExamPaperName:{0}", ExamPaperName);
            //生成试卷
            //Word应用程序变量
            Microsoft.Office.Interop.Word.Application wordexamop = null;
            Document wordexam = null;

            progress = 10;
            backgroundWorker.ReportProgress(progress, "试卷生成中...");

            try
            {
                wordexamop = new Microsoft.Office.Interop.Word.Application();
                wordexamop.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;//屏蔽宏
                wordexamop.Visible = false;//使文档不可见,生成文件过程中文档不可见
                                           //wordexamop.Visible = true;//使文档可见

                Object missing = Missing.Value;//由于使用的是COM库，因此有许多变量需要使用Missing.Value代替

                wordexam = wordexamop.Documents.Add();

                progress = 20;
                backgroundWorker.ReportProgress(progress, "试卷生成中...");

                //纸张大小设置
                if (ExamPaper.ExamPaperTask1.PaperSize == "A3") wordexam.PageSetup.PaperSize = WdPaperSize.wdPaperA3;//设置纸张样式为A3纸
                else if (ExamPaper.ExamPaperTask1.PaperSize == "A4") wordexam.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置纸张样式为A4纸
                else wordexam.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//默认纸张样式为A4规格

                //横竖版面设置
                if (ExamPaper.ExamPaperTask1.Orientation == "竖版") wordexam.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                else if (ExamPaper.ExamPaperTask1.Orientation == "横版") wordexam.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                else wordexam.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

                //页面设置
                wordexam.PageSetup.LeftMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginLeft.ToString()); //2.82CM
                wordexam.PageSetup.RightMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginRight.ToString());
                wordexam.PageSetup.TopMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginTop.ToString());//3.525CM
                wordexam.PageSetup.BottomMargin = float.Parse(ExamPaper.ExamPaperTask1.MarginBottom.ToString());

                //分栏设置
                if (ExamPaper.ExamPaperTask1.TextColumns == "2") wordexam.PageSetup.TextColumns.SetCount(2);

                //添加试卷标题
                wordexamop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordexamop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.MainTitleFontSize);
                wordexamop.Selection.Font.Name = ExamPaper.ExamPaperTask1.MainTitleFont;
                wordexamop.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;//首行缩进为0
                wordexamop.Selection.ParagraphFormat.FirstLineIndent = wordexam.Application.CentimetersToPoints((float)0.75 * 0);//首行缩进为0，全局设置为0
                wordexamop.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                wordexamop.Selection.ParagraphFormat.LineSpacing = float.Parse(ExamPaper.ExamPaperTask1.MainTitleNextLineSpace);
                wordexam.Paragraphs.Last.Range.Text = this.ExamPaper.ExamPaperTask1.TitleName + Environment.NewLine + Environment.NewLine;//试卷标题

                object unite = WdUnits.wdStory;
                wordexamop.Selection.EndKey(ref unite, ref missing);//将光标移动到文本末尾
                wordexamop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.SubTitleFontSize);//次标题大小
                wordexamop.Selection.Font.Name = ExamPaper.ExamPaperTask1.SubTitleFont;
                wordexam.Paragraphs.Last.Range.Text = "单位：___________    姓名：___________    成绩：___________" + Environment.NewLine + Environment.NewLine;

                progress = 39;
                backgroundWorker.ReportProgress(progress, "试卷生成中...");

                int step = 50 / ExamPaper.ExamPaperTask1.QuestionName.Count;

                for (int i = 0; i < ExamPaper.ExamPaperTask1.QuestionName.Count; i++)
                {
                    //题目标题加内容
                    wordexamop.Selection.EndKey(ref unite, ref missing);//将光标移动到文本末尾
                                                                        //wordexamop.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordexamop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    wordexamop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.FirstHeadingTitleFontSize);//四号字体大小
                    wordexamop.Selection.Font.Name = ExamPaper.ExamPaperTask1.FirstHeadingTitleFont;
                    //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值,前面已经设置过了
                    wordexamop.Selection.ParagraphFormat.LineSpacing = float.Parse(ExamPaper.ExamPaperTask1.FirstHeadingTitleNextLineSpace);

                    float sum = 0;
                    for (int j = 0; j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count; j++) sum += ExamPaper.ExamPaperTask1.QuestionGrade[i] * ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].amount;

                    wordexam.Paragraphs.Last.Range.Text = Utilities.IntToBigNum(i + 1) + ExamPaper.ExamPaperTask1.QuestionName[i] + "（每题" + ExamPaper.ExamPaperTask1.QuestionGrade[i] + "分，共 " + sum + "分）" + Environment.NewLine;
                    //i++;
                    wordexamop.Selection.EndKey(ref unite, ref missing);//将光标移动到文本末尾
                    wordexamop.Selection.Font.Size = float.Parse(ExamPaper.ExamPaperTask1.SecondHeadingTitleFontSize);//四号字体大小
                    wordexamop.Selection.Font.Name = ExamPaper.ExamPaperTask1.SecondHeadingTitleFont;
                    //wordexamop.Selection.Font.;
                    //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                    wordexamop.Selection.ParagraphFormat.LineSpacing = float.Parse(ExamPaper.ExamPaperTask1.SecondHeadingTitleNextLineSpace);//固定行距为18

                    //Console.WriteLine("examPaperTask.FirstQuestionList[{0}].Count = {1}", i, examPaperTask.FirstQuestionList[i].Count);
                    for (int j = 0; j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count; j++)
                    {
                        List<List<string>> questionList = new List<List<string>>();
                        questionList = null;
                        while (j < ExamPaper.ExamPaperTask1.FirstQuestionList[i].Count)
                        {
                            List<List<string>> newQuestionList = new List<List<string>>();
                            string questionStoragePath = ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].path;
                            int num = ExamPaper.ExamPaperTask1.FirstQuestionList[i][j].amount;
                            newQuestionList = Utilities.GetNeededQuestionAndAnswerFromQuestionStorage(questionStoragePath, num).Item1;
                            questionList = Utilities.CombineList(questionList, newQuestionList);
                            j++;
                        }

                        for (int k = 0; k < questionList.Count; k++)
                        {
                            //Console.WriteLine("k:{0}", k);
                            for (int l = 0; l < questionList[k].Count; l++)
                            {
                                if (l == 0) wordexam.Paragraphs.Last.Range.Text = (k + 1).ToString() + ". " + questionList[k][l] + Environment.NewLine;
                                else wordexam.Paragraphs.Last.Range.Text = questionList[k][l] + Environment.NewLine;
                            }
                            for (int m = 0; m < ExamPaper.ExamPaperTask1.QuestionNextLineSpace[i]; m++)//每个大的题型的小题行距以每个大题第一个题库输入的行间距为准，后续题库的输入值不予采纳
                                wordexam.Paragraphs.Last.Range.Text = Environment.NewLine;
                        }
                    }
                    wordexam.Paragraphs.Last.Range.Text = Environment.NewLine;//大题之间的空行

                    progress = 39 + i * step;
                    backgroundWorker.ReportProgress(progress, "试卷生成中...");
                }

                progress = 90;
                backgroundWorker.ReportProgress(progress, "试卷已完成");
                Thread.Sleep(1000);

                wordexam.SaveAs(Path.Combine(ExamPaper.ExamPaperDirPath, ExamPaper.ExamPaperName));
                wordexam.Close(ref missing, ref missing, ref missing);
                wordexam = null;

                wordexamop.Quit(ref missing, ref missing, ref missing);
                wordexamop = null;
                progress = 95;
                backgroundWorker.ReportProgress(progress, "正在进行最后步骤...");
                Thread.Sleep(500);
                progress = 99;
                backgroundWorker.ReportProgress(progress, "正在进行最后步骤...");
                Thread.Sleep(500);
                progress = 100;
                backgroundWorker.ReportProgress(progress, "正在进行最后步骤...");
                Thread.Sleep(1000);
            }
            catch (Exception err)
            {
                //Console.WriteLine($"Processing failed: {err.Message}");
                if (progress < 90)
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "试卷生成错误");
                        Thread.Sleep(10);
                    }
                }
                else if (progress == 90)
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "试卷保存错误");
                        Thread.Sleep(10);
                    }
                }
                else
                {
                    for (int i = progress; i > 0; i--)
                    {
                        backgroundWorker.ReportProgress(i, "未可知其错");
                        Thread.Sleep(10);
                    }
                }
                if (err != null) { workResult.Error = err; workResult.Result = err.Message.Replace("\n", "").Replace("\r", "").Replace(Environment.NewLine, ""); e.Cancel = true; }
            }
        }

        #endregion

        /// <summary>
        /// 只生成试卷应用（生成试卷）
        /// </summary>
        /// <param name="examPaper"></param>
        //private void StartLongRunningTask2(ExamPaper examPaper)
        //{
        //    formExamPaperToolStripProgressBar.Visible = true;
        //    //formExamPaperToolStripStatusLabel.Margin.Left = 140;
        //    BackgroundWorker worker = new BackgroundWorker
        //    {
        //        WorkerReportsProgress = true,
        //        WorkerSupportsCancellation = true
        //    };

        //    worker.DoWork += (sender, e) =>
        //    {
        //        // 模拟耗时操作
        //        //for (int i = 0; i < 100; i++)
        //        //{
        //        if (worker.CancellationPending)
        //        {
        //            e.Cancel = true;
        //            return;
        //        }

        //        // 进度更新应该在主线程
        //        worker.ReportProgress(0);
        //        if (examPaper.FormExamPaper(worker)) return;
        //    };

        //    worker.ProgressChanged += (sender, e) =>
        //    {
        //        formExamPaperToolStripProgressBar.Value = e.ProgressPercentage;

        //        if (formExamPaperToolStripProgressBar.Value == 100)
        //        {
        //            this.formExamPaperToolStripStatusLabel.Text = "试卷已生成";
        //            formExamPaperToolStripProgressBar.Visible = false;
        //            return;
        //        }
        //    };

        //    worker.RunWorkerAsync();

        //}

        /// <summary>
        /// 关于窗体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (af == null)
            {
                af = new AboutForm();
                af.StartPosition = FormStartPosition.Manual;
                af.Location = new System.Drawing.Point(this.Location.X, this.Location.Y + 55);
                af.Show();
            }
            else af.Activate();
        }

        /// <summary>
        /// 主窗体退出按钮事件，完全彻底退出应用软件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) { System.Environment.Exit(0); }

        /// <summary>
        /// 使用手册窗体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ManualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (mf == null)
            {
                mf = new ManualForm();
                mf.StartPosition = FormStartPosition.Manual;
                mf.Location = new System.Drawing.Point(this.Location.X, this.Location.Y + 55);
                mf.Show();
            }
            else mf.Activate();
        }

    }
}