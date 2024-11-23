//This file is part of ExamPaper Factory
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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

        public static NewExamTaskContentForm netc = null;//新建任务窗体
        public static EditExamTaskContentForm eetc = null;//编辑任务窗体
        public static ExamTaskContentForm etc = null;//更多内容窗体
        public static AboutForm af = null;//关于内容窗体
        public static ManualForm mf = null;//使用手册窗体窗体

        public bool isCurrentTaskExist = false;//判断当前界面是否有任务，右任务才涉及到任务新旧的问题
        public bool isCurrentTaskNew = false;//是新的任务文件标志（只有编辑过未被保存的文件才算新任务文件）
        public bool isCurrentRight ;//当前任务文件经过检查后，修改此处的值，若处理后发现有问题，则这里为false；处理后没有问题则为true；
        
        public string ExamPaperTaskFilePath { get => examPaperTaskFilePath; set => examPaperTaskFilePath = value; }
        public ExamPaperTask CurrentTask { get => currentTask; set => currentTask = value; }
        public string ExamPaperFilePath { get => examPaperDirPath; set => examPaperDirPath = value; }

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
            if(examPaperTaskFilePath != null) mainFormTitleHere = Path.GetFileName(examPaperTaskFilePath);
            if (!isCurrentTaskNew && examPaperTaskFilePath == null) this.Text = "试卷工厂";
            else if (!isCurrentTaskNew && examPaperTaskFilePath != null) this.Text = "试卷工厂(" + mainFormTitleHere + ")";
            else if (isCurrentTaskNew && examPaperTaskFilePath == null) this.Text = "试卷工厂*" ;
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
                this.saveAsToolStripMenuItem.Enabled =true;//另存为按钮可用
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
                etc.Location = new System.Drawing.Point(this.Location.X,this.Location.Y + 55);
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
                else {
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
            sfd.DefaultExt="xml";//设置默认的文件名
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
        /// 生成试卷按钮（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormExamPaperToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isCurrentRight)
            {
                ExamPaper examPaper = new ExamPaper();//创建试卷生成对象
                examPaper.ExamPaperTask = CurrentTask;//生成试卷所依据的任务文件为当前任务文件
                examPaper.ExamPaperDirPath = examPaperDirPath;//给定试卷生成目录
                examPaper.MainFormTitle = examPaperTaskFilePath != null ? examPaper.MainFormTitle = Path.GetFileNameWithoutExtension(examPaperTaskFilePath) : "";
                examPaper.ExamPaperFileFormTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");//20241113210012(2024年11月13日21:00:12)
                //examPaper.FormExamPaper();
                StartLongRunningTask2(examPaper);
                //examPaper.FormExamPaperAndAnswer();
            }
        }

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

        /// <summary>
        /// 生成试卷和答案按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormExamPaperAndAnswerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isCurrentRight)
            {
                ExamPaper examPaper = new ExamPaper();//创建试卷生成对象
                examPaper.ExamPaperTask = CurrentTask;//生成试卷所依据的任务文件为当前任务文件
                examPaper.ExamPaperDirPath = examPaperDirPath;//给定试卷生成目录
                if (examPaperTaskFilePath != null) examPaper.MainFormTitle = Path.GetFileNameWithoutExtension(examPaperTaskFilePath);
                else examPaper.MainFormTitle = "";
                examPaper.ExamPaperFileFormTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");//20241113210012(2024年11月13日21:00:12)
                //examPaper.FormExamPaper();

                StartLongRunningTask(examPaper);
                //examPaper.FormExamPaperAndAnswer();
            }
        }

        /// <summary>
        /// 生成试卷开辟新线程
        /// </summary>
        /// <param name="examPaper">试卷对象</param>
        private void StartLongRunningTask(ExamPaper examPaper)
        {
            formExamPaperToolStripProgressBar.Visible = true;
            //formExamPaperToolStripStatusLabel.Margin.Left = 140;
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;

            worker.DoWork += (sender, e) =>
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    Console.WriteLine("111");
                    return;
                }

                // 进度更新应该在主线程
                worker.ReportProgress(0);
                if (examPaper.FormExamPaperAndAnswer(worker))
                {
                    Console.WriteLine("113331"); 
                    return;
                }
            };

            worker.ProgressChanged += (sender, e) =>
            {
                formExamPaperToolStripProgressBar.Value = e.ProgressPercentage;

                if (formExamPaperToolStripProgressBar.Value == 100)
                {
                    this.formExamPaperToolStripStatusLabel.Text = "试卷答案已生成";
                    formExamPaperToolStripProgressBar.Visible = false;
                    Console.WriteLine("222");
                    return;
                }
            };

            worker.RunWorkerAsync();
        }

        /// <summary>
        /// 只生成试卷应用
        /// </summary>
        /// <param name="examPaper"></param>
        private void StartLongRunningTask2(ExamPaper examPaper)
        {
            formExamPaperToolStripProgressBar.Visible = true;
            //formExamPaperToolStripStatusLabel.Margin.Left = 140;
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;

            worker.DoWork += (sender, e) =>
            {
                // 模拟耗时操作
                //for (int i = 0; i < 100; i++)
                //{
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                // 进度更新应该在主线程
                worker.ReportProgress(0);
                if (examPaper.FormExamPaper(worker)) return;
            };

            worker.ProgressChanged += (sender, e) =>
            {
                formExamPaperToolStripProgressBar.Value = e.ProgressPercentage;

                if (formExamPaperToolStripProgressBar.Value == 100)
                {
                    this.formExamPaperToolStripStatusLabel.Text = "试卷已生成";
                    formExamPaperToolStripProgressBar.Visible = false;
                    return;
                }
            };

            worker.RunWorkerAsync();

        }

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

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) { System.Environment.Exit(0); }

        private void Button1_Click(object sender, EventArgs e)
        {
            //Console.WriteLine( Utilities.GetScreenScaling(this));
        }

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