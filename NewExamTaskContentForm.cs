//This file is part of ExamPaper Factory
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    public partial class NewExamTaskContentForm : Form
    {
        int firstHeadingNum = 1;//大题数计数器,默认有1题
        int questionStorageNum = 1;//当前大题下题库数计数器，默认当前条件下有1个题库
        private readonly List<int> questionStorageNumList = new List<int>() { 1 };//默认题库列表有一个内容

        //int locationHeight = 32;//GroupBox初始位置
        int locationHeight = 19;//GroupBox初始位置

        private const int FIRST_HEADING_OFFSET = 28;//大题栏偏移量，不小于23，偏大一点可以更好区分
        private const int QUESTION_STORAGE_OFFSET = 21;//题库栏偏移量，最小值为23，不建议改变

        private const string CONTENT_TEXTBOX_FONT = "Times New Roman";//TextBox框字体类型
        private const float CONTENT_TEXTBOX_FONT_SIZE = 8.25f;//TextBox框字体大小

        private const string CONTENT_COMBOBOX_FONT = "Times New Roman";//Combobox框字体类型
        private const float CONTENT_COMBOBOX_FONT_SIZE = 7.75f;//Combobox框字体大小

        float scaling = 1f; //缩放比例(默认无缩放)

        private ExamPaperTask examPaperTask = null;

        private string titleName;//试卷名称

        private string paperSize;//试卷纸张大小
        private string textColumns;//试卷分栏数量
        private string orientation;//试卷排版方向

        private float marginTop;//试卷上边距
        private float marginBottom;//试卷下边距
        private float marginLeft;//试卷左边距
        private float marginRight;//试卷右边距

        //试卷标题相关参数（字体、大小、样式、空行距离）
        private string mainTitleFont;
        private string mainTitleFontSize;
        private string mainTitleFontStyle;
        private string mainTitleNextLineSpace;

        //试卷单位姓名位置相关参数（字体、大小、样式、空行距离）
        private string subTitleFont;
        private string subTitleFontSize;
        private string subTitleFontStyle;
        private string subTitleNextLineSpace;

        //一级标题相关参数（字体、大小、样式、空行距离）
        private string firstHeadingTitleFont;
        private string firstHeadingTitleFontSize;
        private string firstHeadingTitleFontStyle;
        private string firstHeadingTitleNextLineSpace;

        //二级标题相关参数（字体、大小、样式、空行距离）
        private string secondHeadingTitleFont;
        private string secondHeadingTitleFontSize;
        private string secondHeadingTitleFontStyle;
        private string secondHeadingTitleNextLineSpace;

        //小题目单元
        public struct QuestionElement
        {
            private string path;
            private int amount;

            public string Path { get => path; set => path = value; }
            public int Amount { get => amount; set => amount = value; }
        }

        //大题题目列表
        private List<string> questionName;

        //大题下面所有小题的每题分数
        private List<float> questionGrade;

        ////大题下每小题的空行数
        private List<int> questionNextLineSpace;

        //大题完整列表，除了大题名称，包括了里面的所有东西
        private List<List<QuestionElement>> firstQuestionList;

        public string TitleName { get => titleName; set => titleName = value; }
        public string PaperSize { get => paperSize; set => paperSize = value; }
        public string TextColumns { get => textColumns; set => textColumns = value; }
        public string Orientation { get => orientation; set => orientation = value; }
        public float MarginTop { get => marginTop; set => marginTop = value; }
        public float MarginBottom { get => marginBottom; set => marginBottom = value; }
        public float MarginLeft { get => marginLeft; set => marginLeft = value; }
        public float MarginRight { get => marginRight; set => marginRight = value; }
        public string MainTitleFont { get => mainTitleFont; set => mainTitleFont = value; }
        public string MainTitleFontSize { get => mainTitleFontSize; set => mainTitleFontSize = value; }
        public string MainTitleFontStyle { get => mainTitleFontStyle; set => mainTitleFontStyle = value; }
        public string MainTitleNextLineSpace { get => mainTitleNextLineSpace; set => mainTitleNextLineSpace = value; }
        public string SubTitleFont { get => subTitleFont; set => subTitleFont = value; }
        public string SubTitleFontSize { get => subTitleFontSize; set => subTitleFontSize = value; }
        public string SubTitleFontStyle { get => subTitleFontStyle; set => subTitleFontStyle = value; }
        public string SubTitleNextLineSpace { get => subTitleNextLineSpace; set => subTitleNextLineSpace = value; }
        public string FirstHeadingTitleFont { get => firstHeadingTitleFont; set => firstHeadingTitleFont = value; }
        public string FirstHeadingTitleFontSize { get => firstHeadingTitleFontSize; set => firstHeadingTitleFontSize = value; }
        public string FirstHeadingTitleFontStyle { get => firstHeadingTitleFontStyle; set => firstHeadingTitleFontStyle = value; }
        public string FirstHeadingTitleNextLineSpace { get => firstHeadingTitleNextLineSpace; set => firstHeadingTitleNextLineSpace = value; }
        public string SecondHeadingTitleFont { get => secondHeadingTitleFont; set => secondHeadingTitleFont = value; }
        public string SecondHeadingTitleFontSize { get => secondHeadingTitleFontSize; set => secondHeadingTitleFontSize = value; }
        public string SecondHeadingTitleFontStyle { get => secondHeadingTitleFontStyle; set => secondHeadingTitleFontStyle = value; }
        public string SecondHeadingTitleNextLineSpace { get => secondHeadingTitleNextLineSpace; set => secondHeadingTitleNextLineSpace = value; }
        public List<string> QuestionName { get => questionName; set => questionName = value; }
        public List<float> QuestionGrade { get => questionGrade; set => questionGrade = value; }
        internal List<int> QuestionNextLineSpace { get => questionNextLineSpace; set => questionNextLineSpace = value; }
        public List<List<QuestionElement>> FirstQuestionList { get => firstQuestionList; set => firstQuestionList = value; }
        public float Scaling { get => scaling; set => scaling = value; }


        /// <summary>
        /// 【3】定义委托变量
        /// </summary>
        public CurrentTaskDelegate currentTaskDelegate;

        /// <summary>
        /// 构造函数，初始化方法
        /// </summary>
        public NewExamTaskContentForm()
        {
            Scaling = Utilities.GetScreenScaling(form: this);
            locationHeight = (int)(locationHeight * Scaling);
            InitializeComponent();
        }

        /// <summary>
        /// 取消按钮，直接推出新建文件框体，不做任何操作（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CancelToolStripMenuItem_Click(object sender, EventArgs e) { this.Close(); }

        /// <summary>
        /// 题库-方法的重载（已更新）
        /// </summary>
        private void QuestionStorageMinus()
        {
            int lastQuestionStorageNum = questionStorageNumList[questionStorageNumList.Count - 1];//提取最后一位数
            if (lastQuestionStorageNum == 1) return;//没有多余题库的情况直接不反应

            //删除最后一条题库栏
            foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
            { if (tb.Name == "pathLabel" + firstHeadingNum.ToString() + "_" + lastQuestionStorageNum.ToString()) contentPanel.Controls.Remove(tb); }
            foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
            { if (cbb.Name == "amountComboBox" + firstHeadingNum.ToString() + "_" + lastQuestionStorageNum.ToString()) contentPanel.Controls.Remove(cbb); }


            locationHeight -= QUESTION_STORAGE_OFFSET;//位置变更
            contentPanel.Height -= QUESTION_STORAGE_OFFSET;//contentPanel组件大小变更
            examPaperContentGroupBox.Height -= QUESTION_STORAGE_OFFSET;//GroupBox组件大小变更
            Height -= QUESTION_STORAGE_OFFSET;//框体大小变更
            questionStorageNumList[questionStorageNumList.Count - 1]--;//questionStorageNumList列表的最后一位数减1
            questionStorageNum--;//最后一个大题的题库数减1
        }

        /// <summary>
        /// 题型-方法的重载（已更新）
        /// </summary>
        private void FirstHeadingMinus()
        {
            int lastQuestionStorageNum = questionStorageNumList[questionStorageNumList.Count - 1];//提取题库数列表最后一位数

            //测试用
            //foreach(System.Windows.Forms.Label l in examPaperContentGroupBox.Controls.OfType<System.Windows.Forms.Label>())
            //{ Console.WriteLine("l.Name:{0}", l.Name); }

            if (firstHeadingNum == 1) return;//如果仅剩最后一条，则不做操作
            //删掉大题开头的一条TextBox题目框
            foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
            { if (tb.Name == "firstHeadingTextBox" + firstHeadingNum.ToString()) contentPanel.Controls.Remove(tb); }
            //删掉大题末尾的一条gradeComboBox选择框
            foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
            { if (cb.Name == "gradeComboBox" + firstHeadingNum.ToString()) contentPanel.Controls.Remove(cb); }
            //删掉大题末尾的一条nextLineSpaceComboBox选择框
            foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
            { if (cb.Name == "nextLineSpaceComboBox" + firstHeadingNum.ToString()) contentPanel.Controls.Remove(cb); }

            //循环删除该大题下的所有题库栏，包括大题本体所带的题库栏
            for (int i = 0; i < lastQuestionStorageNum; i++)
            {
                //Console.WriteLine("i:{0}", i);//测试用
                foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
                { if (tb.Name == "pathTextBox" + firstHeadingNum.ToString() + "_" + (i + 1).ToString()) contentPanel.Controls.Remove(tb); }
                foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                { if (cbb.Name == "amountComboBox" + firstHeadingNum.ToString() + "_" + (i + 1).ToString()) contentPanel.Controls.Remove(cbb); }
            }

            locationHeight = locationHeight - (lastQuestionStorageNum - 1) * QUESTION_STORAGE_OFFSET - FIRST_HEADING_OFFSET;//位置往回收缩量
            contentPanel.Height -= (QUESTION_STORAGE_OFFSET * (lastQuestionStorageNum - 1) + FIRST_HEADING_OFFSET);
            examPaperContentGroupBox.Height = examPaperContentGroupBox.Height - QUESTION_STORAGE_OFFSET * (lastQuestionStorageNum - 1) - FIRST_HEADING_OFFSET;
            Height = Height - QUESTION_STORAGE_OFFSET * (lastQuestionStorageNum - 1) - FIRST_HEADING_OFFSET;//框体大小收缩量

            firstHeadingNum--;//大题数目减1
            questionStorageNumList.RemoveAt(questionStorageNumList.Count - 1);//删除questionStorageNumList列表最后一位数值
            questionStorageNum = questionStorageNumList[questionStorageNumList.Count - 1];//重新赋值为questionStorageNumList删掉最后一位数值后的最后一位
            //Console.WriteLine("题库- 结束之后 firstHeadingNum:{0}", firstHeadingNum);//测试用
        }

        /// <summary>
        /// 题型+按钮，在当前界面新增一条填写大题信息的内容栏，
        /// 改变Groupbox和整个窗体的大小（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FirstHeadingPlusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (firstHeadingNum == 19) return;//最多不超过19个大题,一般情况下用不上那么多，但还是有个限制
            //locationHeight = (int)(locationHeight * scaling);

            //Console.WriteLine(Height);
            //Console.WriteLine("题库+ 开始时 firstHeadingNum:{0}", firstHeadingNum);
            Height += (int)(FIRST_HEADING_OFFSET);//程序外界面框也向下增加25的长度
            examPaperContentGroupBox.Height += (int)(FIRST_HEADING_OFFSET);//GroupBox框增加25的长度
            contentPanel.Height += FIRST_HEADING_OFFSET;//ContentPanel的下边界增加一个偏移量

            locationHeight += (int)(FIRST_HEADING_OFFSET);//新的数据起始点位置向下移25

            questionStorageNum = 1;//新建一个新的题型后，题库数计数复位为1，表明新建的一条题型最少有一个题库
            firstHeadingNum++;//大题计数器增加1
            //Console.WriteLine(Height);
            //如果题库列表数量小于大题数，则列表添加一项，为1
            if (questionStorageNumList.Count < firstHeadingNum) { questionStorageNumList.Add(questionStorageNum); }

            //firstHeadingTextBox新增的题型名称输入框
            System.Windows.Forms.TextBox firstHeadingName = new System.Windows.Forms.TextBox();
            firstHeadingName.Name = "firstHeadingTextBox" + firstHeadingNum.ToString();
            firstHeadingName.Location = new Point((int)(2 * Scaling), locationHeight);
            firstHeadingName.Size = new Size((int)(80 * Scaling), (int)(20 * Scaling));
            firstHeadingName.Text = "";
            firstHeadingName.Font = new Font(CONTENT_TEXTBOX_FONT, CONTENT_TEXTBOX_FONT_SIZE);
            firstHeadingName.BorderStyle = BorderStyle.Fixed3D;
            firstHeadingName.TextAlign = HorizontalAlignment.Left;
            firstHeadingName.BackColor = SystemColors.Control;
            firstHeadingName.TextChanged += new System.EventHandler(this.IsValid_TextChanged);

            //firstHeadingPathTextBox
            System.Windows.Forms.TextBox path = new System.Windows.Forms.TextBox();
            path.Name = "pathTextBox" + firstHeadingNum.ToString() + "_" + questionStorageNum.ToString();
            path.Location = new Point((int)(83 * Scaling), locationHeight);
            path.Size = new Size((int)(283 * Scaling), (int)(20 * Scaling));
            path.Text = "";
            path.Font = new Font(CONTENT_TEXTBOX_FONT, CONTENT_TEXTBOX_FONT_SIZE);
            path.BorderStyle = BorderStyle.Fixed3D;
            path.TextAlign = HorizontalAlignment.Left;
            path.BackColor = SystemColors.Control;
            path.ReadOnly = true;
            path.TextChanged += new System.EventHandler(this.IsValid_TextChanged);
            path.DoubleClick += new EventHandler(this.PathTextBox_DoubleClick);

            //firstHeadingAmountComobox 题目数量
            System.Windows.Forms.ComboBox amountCombobox = new System.Windows.Forms.ComboBox();
            amountCombobox.Name = "amountComboBox" + firstHeadingNum.ToString() + "_" + questionStorageNum.ToString();
            amountCombobox.Location = new Point((int)(366 * Scaling), (int)(locationHeight));
            amountCombobox.Size = new Size((int)(50 * Scaling), (int)(20 * Scaling));
            amountCombobox.Font = new Font(CONTENT_COMBOBOX_FONT, CONTENT_COMBOBOX_FONT_SIZE);
            amountCombobox.BackColor = SystemColors.Control;
            amountCombobox.DropDownStyle = ComboBoxStyle.DropDownList;
            amountCombobox.FormattingEnabled = true;
            amountCombobox.Enabled = false;
            amountCombobox.Items.AddRange(Utilities.NumToStringList(100));
            amountCombobox.TextChanged += new System.EventHandler(this.IsValid_TextChanged);

            //firstHeadingGradeComobox 单个题目分数
            System.Windows.Forms.ComboBox gradeCombobox = new System.Windows.Forms.ComboBox();
            gradeCombobox.Name = "gradeComboBox" + firstHeadingNum.ToString();
            gradeCombobox.Location = new Point((int)(416 * Scaling), (int)(locationHeight));
            gradeCombobox.Size = new Size((int)(50 * Scaling), (int)(20 * Scaling));
            gradeCombobox.Font = new Font(CONTENT_COMBOBOX_FONT, CONTENT_COMBOBOX_FONT_SIZE);
            gradeCombobox.BackColor = SystemColors.Control;
            gradeCombobox.DropDownStyle = ComboBoxStyle.DropDownList;
            gradeCombobox.FormattingEnabled = true;
            gradeCombobox.Items.AddRange(Utilities.NumToStringList(0.5f, 100, 0.5f));
            gradeCombobox.TextChanged += new System.EventHandler(this.IsValid_TextChanged);

            //firstHeadingNextLineSpaceCombobox 小题之间的行间距
            System.Windows.Forms.ComboBox nextLineSpaceCombobox = new System.Windows.Forms.ComboBox();
            nextLineSpaceCombobox.Name = "nextLineSpaceComboBox" + firstHeadingNum.ToString();
            nextLineSpaceCombobox.Location = new Point((int)(466 * Scaling), (int)(locationHeight));
            nextLineSpaceCombobox.Size = new Size((int)(50 * Scaling), (int)(20 * Scaling));
            nextLineSpaceCombobox.Font = new Font(CONTENT_COMBOBOX_FONT, CONTENT_COMBOBOX_FONT_SIZE);
            nextLineSpaceCombobox.BackColor = SystemColors.Control;
            nextLineSpaceCombobox.DropDownStyle = ComboBoxStyle.DropDownList;
            nextLineSpaceCombobox.FormattingEnabled = true;
            nextLineSpaceCombobox.Items.AddRange(Utilities.NumToStringList(0, 21));
            nextLineSpaceCombobox.TextChanged += new System.EventHandler(this.IsValid_TextChanged);

            //一同添加进contentPanel控件
            contentPanel.Controls.Add(firstHeadingName);
            contentPanel.Controls.Add(path);
            contentPanel.Controls.Add(amountCombobox);
            contentPanel.Controls.Add(gradeCombobox);
            contentPanel.Controls.Add(nextLineSpaceCombobox);

            IsDataValid();//完事后，进行合法性检查
        }

        /// <summary>
        /// 双击选择路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PathTextBox_DoubleClick(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox textBoxButton = sender as System.Windows.Forms.TextBox;
            OpenFileDialog ofd = new OpenFileDialog();//打开文件
            ofd.Filter = "(*.txt)|*.txt";//设置过滤器，仅能选择txt文件类型
            //ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //默认获取“我的电脑”路径
            ofd.Title = "请选择题库文件";//文件选择对话框标题
            ofd.Multiselect = false; //不允许多重选择
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBoxButton.Text = ofd.FileName;
                textBoxButton.SelectionStart = textBoxButton.Text.Length;
            }
            ShouldAmountButtonEnabled(textBoxButton.Name.Substring(textBoxButton.Name.Length - 3));
        }

        /// <summary>
        /// 按下题型-按钮后，在当前界面上减去一整个大题信息的内容栏，
        /// 并同时改变Groupbox和整个窗体的大小（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FirstHeadingMinusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int lastQuestionStorageNum = questionStorageNumList[questionStorageNumList.Count - 1];//提取题库数列表最后一位数

            //测试用
            //foreach(System.Windows.Forms.Label l in examPaperContentGroupBox.Controls.OfType<System.Windows.Forms.Label>())
            //{ Console.WriteLine("l.Name:{0}", l.Name); }

            if (firstHeadingNum == 1) return;//如果仅剩最后一条，则不做操作
            //删掉大题开头的一条TextBox题目框
            foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
            { if (tb.Name == "firstHeadingTextBox" + firstHeadingNum.ToString()) contentPanel.Controls.Remove(tb); }
            //删掉大题末尾的一条gradeComboBox选择框
            foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
            { if (cb.Name == "gradeComboBox" + firstHeadingNum.ToString()) contentPanel.Controls.Remove(cb); }
            //删掉大题末尾的一条nextLineSpaceComboBox选择框
            foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
            { if (cb.Name == "nextLineSpaceComboBox" + firstHeadingNum.ToString()) contentPanel.Controls.Remove(cb); }

            //循环删除该大题下的所有题库栏，包括大题本体所带的题库栏
            for (int i = 0; i < lastQuestionStorageNum; i++)
            {
                //Console.WriteLine("i:{0}", i);//测试用
                foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
                { if (tb.Name == "pathTextBox" + firstHeadingNum.ToString() + "_" + (i + 1).ToString()) contentPanel.Controls.Remove(tb); }
                foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                { if (cbb.Name == "amountComboBox" + firstHeadingNum.ToString() + "_" + (i + 1).ToString()) contentPanel.Controls.Remove(cbb); }

            }

            locationHeight = locationHeight - (lastQuestionStorageNum - 1) * QUESTION_STORAGE_OFFSET - FIRST_HEADING_OFFSET;//位置往回收缩量
            contentPanel.Height -= (QUESTION_STORAGE_OFFSET * (lastQuestionStorageNum - 1) + FIRST_HEADING_OFFSET);//contentPanel下边界回缩量
            examPaperContentGroupBox.Height = examPaperContentGroupBox.Height - QUESTION_STORAGE_OFFSET * (lastQuestionStorageNum - 1) - FIRST_HEADING_OFFSET;//examPaperContentGroupBox下边界回缩量
            Height = Height - QUESTION_STORAGE_OFFSET * (lastQuestionStorageNum - 1) - FIRST_HEADING_OFFSET;//框体大小收缩量
            firstHeadingNum--;//大题数目减1
            questionStorageNumList.RemoveAt(questionStorageNumList.Count - 1);//删除questionStorageNumList列表最后一位数值
            questionStorageNum = questionStorageNumList[questionStorageNumList.Count - 1];//重新赋值为questionStorageNumList删掉最后一位数值后的最后一位

            IsDataValid();//完事后，进行合法性检查
                          //Console.WriteLine("题库- 结束之后 firstHeadingNum:{0}", firstHeadingNum);//测试用

            //Console.WriteLine("题型- 结束之后: ");
            //Console.WriteLine("大题数firstHeadingNum:{0}", firstHeadingNum);
            //Console.WriteLine("最后大题的题库数questionStorageNum:{0}", questionStorageNum);
            //Console.WriteLine("链表里大题数questionStorageNumList.Count:{0}", questionStorageNumList.Count);
            //Console.WriteLine("链表里最后大题题库数questionStorageNumList.LastNum:{0}", questionStorageNumList[questionStorageNumList.Count - 1]);
            //for (int i = 0; i < questionStorageNumList.Count; i++)
            //{ Console.WriteLine("questionStorageList[{0}] = {1}", i, questionStorageNumList[i]); }
        }

        /// <summary>
        /// 题库+按钮按下后的效果（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void QuestionStoragePlusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (questionStorageNum == 9) return;//单个题型最多不超过9个题库路径，已经很疯狂了
            //Console.WriteLine("题库+ 开始时 questionStorageNum:{0}", questionStorageNum);//测试用
            questionStorageNum++;//添加一条题库记录

            //对列表进行数据更新，主要是添加或最后的一个值加1
            if (questionStorageNumList.Count < firstHeadingNum) { questionStorageNumList.Add(questionStorageNum); }
            else { questionStorageNumList[questionStorageNumList.Count - 1] = questionStorageNum; }

            Height += (int)(QUESTION_STORAGE_OFFSET);//程序外界面框向下增加20的长度
            examPaperContentGroupBox.Height += (int)QUESTION_STORAGE_OFFSET;//examPaperContentGroupBox框向下增加20长度
            contentPanel.Height += (int)QUESTION_STORAGE_OFFSET;//contentPanel框向下增加20长度

            locationHeight += QUESTION_STORAGE_OFFSET;//新的数据其实点位置向下移23，因为题库路径需要紧贴第一条的题型，逻辑上更好进行区分

            //firstHeadingPathTextBox 路径地址栏选择
            System.Windows.Forms.TextBox path = new System.Windows.Forms.TextBox();
            path.Name = "pathTextBox" + firstHeadingNum.ToString() + "_" + questionStorageNum.ToString();
            path.Location = new Point((int)(83 * Scaling), locationHeight);
            path.Size = new Size((int)(283 * Scaling), (int)(20 * Scaling));
            path.Text = "";
            path.Font = new Font(CONTENT_TEXTBOX_FONT, CONTENT_TEXTBOX_FONT_SIZE);
            path.BorderStyle = BorderStyle.Fixed3D;
            path.TextAlign = HorizontalAlignment.Left;
            path.BackColor = SystemColors.Control;
            path.ReadOnly = true;
            path.TextChanged += new EventHandler(this.IsValid_TextChanged);
            path.DoubleClick += new EventHandler(this.PathTextBox_DoubleClick);

            //firstHeadingAmountComobox 题目数量
            System.Windows.Forms.ComboBox amountCombobox = new System.Windows.Forms.ComboBox();
            amountCombobox.Name = "amountComboBox" + firstHeadingNum.ToString() + "_" + questionStorageNum.ToString();
            amountCombobox.Location = new Point((int)(366 * Scaling), (int)(locationHeight));
            amountCombobox.Font = new Font(CONTENT_COMBOBOX_FONT, CONTENT_COMBOBOX_FONT_SIZE);
            amountCombobox.Size = new Size((int)(50 * Scaling), (int)(20 * Scaling));
            amountCombobox.BackColor = SystemColors.Control;
            amountCombobox.DropDownStyle = ComboBoxStyle.DropDownList;
            amountCombobox.FormattingEnabled = true;
            amountCombobox.Enabled = false;
            amountCombobox.Items.AddRange(Utilities.NumToStringList(100));
            amountCombobox.TextChanged += new EventHandler(this.IsValid_TextChanged);

            //一同添加进contentPanel控件
            contentPanel.Controls.Add(path);
            contentPanel.Controls.Add(amountCombobox);

            IsDataValid();//完事后，进行合法性检查
        }

        /// <summary>
        /// 题库-按钮作用（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void QuestionStorageMinusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int lastQuestionStorageNum = questionStorageNumList[questionStorageNumList.Count - 1];//提取最后一位数
            if (lastQuestionStorageNum == 1) return;//没有多余题库的情况直接不反应

            //删除最后一条题库栏
            foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
            { if (tb.Name == "pathTextBox" + firstHeadingNum.ToString() + "_" + lastQuestionStorageNum.ToString()) contentPanel.Controls.Remove(tb); }

            foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
            { if (cbb.Name == "amountComboBox" + firstHeadingNum.ToString() + "_" + lastQuestionStorageNum.ToString()) contentPanel.Controls.Remove(cbb); }

            contentPanel.Height -= QUESTION_STORAGE_OFFSET;//contentPanel组件大小变更
            examPaperContentGroupBox.Height -= QUESTION_STORAGE_OFFSET;//GroupBox组件大小变更
            Height -= QUESTION_STORAGE_OFFSET;//框体大小变更
            locationHeight -= QUESTION_STORAGE_OFFSET;//位置变更
            questionStorageNumList[questionStorageNumList.Count - 1]--;//questionStorageNumList列表的最后一位数减1
            questionStorageNum--;//最后一个大题的题库数减1

            IsDataValid();//完事后，进行合法性检查

            //Console.WriteLine("题库- 结束之后: ");
            //Console.WriteLine("大题数firstHeadingNum:{0}", firstHeadingNum);
            //Console.WriteLine("最后大题的题库数questionStorageNum:{0}", questionStorageNum);
            //Console.WriteLine("链表里大题数questionStorageNumList.Count:{0}", questionStorageNumList.Count);
            //Console.WriteLine("链表里最后大题题库数questionStorageNumList.LastNum:{0}", questionStorageNumList[questionStorageNumList.Count - 1]);
            //for (int i = 0; i < questionStorageNumList.Count; i++)
            //{ Console.WriteLine("questionStorageList[{0}] = {1}", i, questionStorageNumList[i]); }
        }

        /// <summary>
        /// 确定按钮，按下后将所有输入的数据进行合法化检查，
        /// 不合法的数据一律进行无害化处理（取默认值）
        /// 将数据传递给格式化数据类ExamPaperTask（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void YesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TitleName = MainTitleTextBox.Text;

            PaperSize = paperSizeComboBox.Text;
            TextColumns = textColumnsComboBox.Text;
            Orientation = orientationComboBox.Text;

            MarginTop = float.Parse(marginTopComboBox.Text);
            MarginBottom = float.Parse(marginBottomComboBox.Text);
            MarginLeft = float.Parse(marginLeftComboBox.Text);
            MarginRight = float.Parse(marginRightComboBox.Text);

            MainTitleFont = mainTitleFontLabel.Text;
            MainTitleFontSize = mainTitleFontSizeLabel.Text;
            MainTitleFontStyle = mainTitleFontStyleLabel.Text;
            MainTitleNextLineSpace = mainTitleNextSpacingLineComboBox.Text;

            SubTitleFont = subTitleFontLabel.Text;
            SubTitleFontSize = subTitleFontSizeLabel.Text;
            SubTitleFontStyle = subTitleFontStyleLabel.Text;
            SubTitleNextLineSpace = subTitleNextSpacingLineComboBox.Text;

            FirstHeadingTitleFont = firstHeadingFontLabel.Text;
            FirstHeadingTitleFontSize = firstHeadingFontSizeLabel.Text;
            FirstHeadingTitleFontStyle = firstHeadingFontStyleLabel.Text;
            FirstHeadingTitleNextLineSpace = firstHeadingNextSpacingLineComboBox.Text;

            SecondHeadingTitleFont = secondHeadingFontLabel.Text;
            SecondHeadingTitleFontSize = secondHeadingFontSizeLabel.Text;
            SecondHeadingTitleFontStyle = secondHeadingFontStyleLabel.Text;
            SecondHeadingTitleNextLineSpace = secondHeadingNextSpacingLineComboBox.Text;

            FirstQuestionList = new List<List<QuestionElement>>();

            QuestionName = new List<string>();
            QuestionGrade = new List<float>();
            QuestionNextLineSpace = new List<int>();


            for (int i = 1; i <= firstHeadingNum; i++)
            {
                foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
                { if (tb.Name == "firstHeadingTextBox" + i.ToString()) QuestionName.Add(tb.Text); }
                foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                { if (cb.Name == "gradeComboBox" + i.ToString()) QuestionGrade.Add(float.Parse(cb.Text)); }
                foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                { if (cb.Name == "nextLineSpaceComboBox" + i.ToString()) QuestionNextLineSpace.Add(int.Parse(cb.Text)); }

                List<QuestionElement> qeList = new List<QuestionElement>();
                for (int j = 0; j < questionStorageNumList[i - 1]; j++)
                {
                    QuestionElement qe = new QuestionElement();
                    foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
                    { if (tb.Name == "pathTextBox" + i.ToString() + "_" + (j + 1).ToString()) qe.Path = tb.Text; }
                    foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                    { if (cbb.Name == "amountComboBox" + i.ToString() + "_" + (j + 1).ToString()) { qe.Amount = int.Parse(cbb.Text); } }
                    qeList.Add(qe);
                }
                FirstQuestionList.Add(qeList);
            }
            //for (int i = 0; i < FirstQuestionList.Count; i++)
            //{
            //    Console.WriteLine("第{0}题:{1}", 1 + i, QuestionName[i]);
            //    for (int j = 0; j < FirstQuestionList[i].Count; j++)
            //    {
            //        string pathname = FirstQuestionList[i][j].Path;
            //        double gradename = FirstQuestionList[i][j].Grade;
            //        double amountname = FirstQuestionList[i][j].Amount;
            //        Console.WriteLine("question[{0}][{1}]: path:{2}, grade:{3}, amount:{4}", i, j, pathname, gradename, amountname);
            //    }
            //}

            //for (int i = 0; i < QuestionName.Count; i++)
            //{ Console.WriteLine(questionName[i]); }

            examPaperTask = new ExamPaperTask(); //新建一个ExamPaperTask变量
            examPaperTask.ReadNewExamTaskContentFormData(this);//调用ExamPaperTask方法将此信息框的所有信息对新变量赋值
            //使用委托变量将数据传回到主界面
            currentTaskDelegate(examPaperTask);
            this.Close();
        }

        /// <summary>
        /// 数据完整性检查，在每进行一次数据变更时调用该方法，
        /// 该方法检查所有的数据均合法后再将确认键变更为可用状态（已更新）
        /// </summary>
        /// <returns></returns>
        private void IsDataValid()
        {
            //Console.WriteLine("1");
            if (this.MainTitleTextBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }//标题为空，退出

            //判断页边距是否为空，因为只能进行选择，不判断数据内容是否合法
            if (this.marginTopComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.marginBottomComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.marginLeftComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.marginRightComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }

            //判断版面内容是否为空
            if (this.orientationComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.textColumnsComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.paperSizeComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }

            //判断试卷标题属性内容是否为空
            if (this.mainTitleFontLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.mainTitleFontStyleLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.mainTitleFontSizeLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.mainTitleNextSpacingLineComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }

            //判断单位名称属性内容是否为空
            if (this.subTitleFontLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.subTitleFontStyleLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.subTitleFontSizeLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.subTitleNextSpacingLineComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }

            //判断一级标题属性内容是否为空
            if (this.firstHeadingFontLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.firstHeadingFontStyleLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.firstHeadingFontSizeLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.firstHeadingNextSpacingLineComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }

            //判断二级标题属性内容是否为空
            if (this.secondHeadingFontLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.secondHeadingFontStyleLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.secondHeadingFontSizeLabel.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }
            if (this.secondHeadingNextSpacingLineComboBox.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; }

            //Console.WriteLine("2 ");
            //考试内容方面
            for (int i = 1; i < questionStorageNumList.Count + 1; i++)
            {
                foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
                { if (tb.Name == "firstHeadingTextBox" + i.ToString() && tb.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; } }
                foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                { if (cb.Name == "gradeComboBox" + i.ToString() && cb.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; } }
                foreach (System.Windows.Forms.ComboBox cb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                { if (cb.Name == "nextLineSpaceComboBox" + i.ToString() && cb.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; } }

                //Console.WriteLine("3.{0}", i);
                for (int j = 1; j < questionStorageNumList[i - 1] + 1; j++)
                {
                    //Console.WriteLine("i:{0}", i);//测试用
                    foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
                    { if (tb.Name == "pathTextBox" + i.ToString() + "_" + j.ToString() && tb.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; } }
                    foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                    { if (cbb.Name == "amountComboBox" + i.ToString() + "_" + j.ToString() && cbb.Text == "") { this.yesToolStripMenuItem.Enabled = false; return; } }

                    //Console.WriteLine("4");
                }

            }
            this.yesToolStripMenuItem.Enabled = true;//全部检查完成后没有问题，"确认"键可使用
            return;
        }

        /// <summary>
        /// 判断并改变题数选择框的可用性,使用固定位置tail角标
        /// </summary>
        /// <param name="tail">角标（形如：1_1）</param>
        private void ShouldAmountButtonEnabled(string tail)
        {
            //Console.WriteLine("i,j:{0},{1}", i, j);//测试用
            foreach (System.Windows.Forms.TextBox tb in contentPanel.Controls.OfType<System.Windows.Forms.TextBox>())
            {
                if (tb.Name == "pathTextBox" + tail)
                {
                    //Console.WriteLine("pathTextBox{0}.Text = {1}", tail, tb.Text);//题库路径
                    if (Utilities.CheckQuestionStorageFormat(tb.Text))
                    {
                        int num = Utilities.CountQuestionNumFromQuestionStoragePath(tb.Text);
                        //Console.WriteLine("pathTextBox{0}.num = {1}", tail, num);//题库对应的题目总数
                        //Console.WriteLine();
                        foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                        {
                            if (cbb.Name == "amountComboBox" + tail)
                            {
                                int currentSelectedIndex = cbb.SelectedIndex;//旧值
                                //Console.WriteLine("amountComboBox{0}.OldSelectedIndex = {1}", tail, currentSelectedIndex);
                                cbb.Items.Clear();
                                cbb.Items.AddRange(Utilities.NumToStringList(num));//设置数量范围（1:num）
                                //Console.WriteLine("num-1: {0},  currentSelectedIndex: {1}", num - 1, currentSelectedIndex);
                                cbb.SelectedIndex = num - 1 > currentSelectedIndex ? currentSelectedIndex : num - 1;
                                cbb.Enabled = true;
                            }
                        }
                    }
                    else
                    {
                        foreach (System.Windows.Forms.ComboBox cbb in contentPanel.Controls.OfType<System.Windows.Forms.ComboBox>())
                        {
                            if (cbb.Name == "amountComboBox" + tail)
                            {
                                cbb.Items.Clear();
                                cbb.Items.AddRange(new string[] { "" });
                                cbb.SelectedItem = "";
                                cbb.Enabled = false;
                            }
                        }
                    }
                }
            }
        }

        #region 窗体所有值改变后进行的操作（已更新）
        private void MainTitleTextBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void MarginTopComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void MarginBottomComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void MarginLeftComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void MarginRightComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        private void ExamPaperSizeComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void TextColumnsComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void OrientationComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        private void MainTitleFontLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void MainTitleFontStyleLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void MainTitleFontSizeLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void MainTitleNextSpacingLineComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        private void SubTitleFontLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void SubTitleFontStyleLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void SubTitleFontSizeLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void SubTitleNextSpacingLineComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        private void FirstHeadingFontLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void FirstHeadingFontStyleLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void FirstHeadingFontSizeLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void FirstHeadingNextSpacingLineComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        private void SecondHeadingFontLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void SecondHeadingFontStyleLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void SecondHeadingFontSizeLabel_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void SecondHeadingNextSpacingLineComboBox_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        private void FirstHeadingTextBox1_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void PathTextBox1_1_TextChanged(object sender, EventArgs e) { this.IsDataValid(); ShouldAmountButtonEnabled("1_1"); }
        private void AmountComboBox1_1_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void GradeComboBox1_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }
        private void NextLineSpaceComboBox1_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        /// <summary>
        /// 新建项目的值变更后进行的方法，主要也就是进行合法性检查（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IsValid_TextChanged(object sender, EventArgs e) { this.IsDataValid(); }

        #endregion

        /// <summary>
        /// 新建任务界面关闭时，同时关闭其对象（不变）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewExamTaskContentForm_FormClosing(object sender, FormClosingEventArgs e) { MainForm.netc = null; }

        /// <summary>
        /// 全复位（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AllRenewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.MainTitleTextBox.Text = "**考试";
            this.marginTopComboBox.Text = "100";
            this.marginBottomComboBox.Text = "80";
            this.marginLeftComboBox.Text = "80";
            this.marginRightComboBox.Text = "80";

            this.paperSizeComboBox.Text = "A4";
            this.textColumnsComboBox.Text = "1";
            this.orientationComboBox.Text = "竖版";

            this.mainTitleFontLabel.Text = "方正小标宋简体";
            this.subTitleFontLabel.Text = "黑体";
            this.firstHeadingFontLabel.Text = "黑体";
            this.secondHeadingFontLabel.Text = "宋体";

            this.mainTitleFontStyleLabel.Text = "Regular";
            this.subTitleFontStyleLabel.Text = "Regular";
            this.firstHeadingFontStyleLabel.Text = "Regular";
            this.secondHeadingFontStyleLabel.Text = "Regular";

            this.mainTitleFontSizeLabel.Text = "20";
            this.subTitleFontSizeLabel.Text = "14";
            this.firstHeadingFontSizeLabel.Text = "14";
            this.secondHeadingFontSizeLabel.Text = "12";

            this.mainTitleNextSpacingLineComboBox.Text = "26";
            this.subTitleNextSpacingLineComboBox.Text = "14";
            this.firstHeadingNextSpacingLineComboBox.Text = "20";
            this.secondHeadingNextSpacingLineComboBox.Text = "18";

            this.firstHeadingTextBox1.Text = "请输入题目";
            this.pathTextBox1_1.Text = "双击选择题库路径";
            this.amountComboBox1_1.Text = "10";
            this.gradeComboBox1.Text = "1.5";
            this.nextLineSpaceComboBox1.Text = "0";

            int firstHeadingNum_Here = firstHeadingNum;
            int questionStorageNum_Here = questionStorageNum;
            for (int i = 0; i < firstHeadingNum_Here; i++) { FirstHeadingMinus(); }//循环删除大题
            for (int i = 0; i < questionStorageNum_Here; i++) { QuestionStorageMinus(); }//循环删除题库
            firstHeadingNum = 1;
            questionStorageNum = 1;
        }

        /// <summary>
        /// 仅复位布局（已更新）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormatRenewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.marginTopComboBox.Text = "100";
            this.marginBottomComboBox.Text = "80";
            this.marginLeftComboBox.Text = "80";
            this.marginRightComboBox.Text = "80";

            this.paperSizeComboBox.Text = "A4";
            this.textColumnsComboBox.Text = "1";
            this.orientationComboBox.Text = "竖版";

            this.mainTitleFontLabel.Text = "方正小标宋简体";
            this.subTitleFontLabel.Text = "黑体";
            this.firstHeadingFontLabel.Text = "黑体";
            this.secondHeadingFontLabel.Text = "宋体";

            this.mainTitleFontStyleLabel.Text = "Regular";
            this.subTitleFontStyleLabel.Text = "Regular";
            this.firstHeadingFontStyleLabel.Text = "Regular";
            this.secondHeadingFontStyleLabel.Text = "Regular";

            this.mainTitleFontSizeLabel.Text = "20";
            this.subTitleFontSizeLabel.Text = "14";
            this.firstHeadingFontSizeLabel.Text = "14";
            this.secondHeadingFontSizeLabel.Text = "12";

            this.mainTitleNextSpacingLineComboBox.Text = "26";
            this.subTitleNextSpacingLineComboBox.Text = "14";
            this.firstHeadingNextSpacingLineComboBox.Text = "20";
            this.secondHeadingNextSpacingLineComboBox.Text = "18";
        }

        #region 试卷各级标题格式按钮动作和功能集合

        /// <summary>
        /// 鼠标单击按下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainTitleFontButtonLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }
        /// <summary>
        /// 鼠标离开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainTitleFontButtonLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }
        /// <summary>
        /// 鼠标放在上面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainTitleFontButtonLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }
        /// <summary>
        /// 主标题格式选择按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainTitleFontButtonLabel_Click(object sender, EventArgs e)
        {
            // 创建字体选择对话框
            FontDialog fontDialog = new FontDialog
            {
                AllowScriptChange = false,
                ShowColor = false,
                ShowHelp = false,
                ShowEffects = false
            };

            // 处理用户的字体选择
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取用户所选字体
                Font selectedFont = fontDialog.Font;

                // 在 label1 中显示所选字体的样式
                this.mainTitleFontLabel.Text = selectedFont.Name;
                this.mainTitleFontSizeLabel.Text = selectedFont.Size.ToString();
                this.mainTitleFontStyleLabel.Text = selectedFont.Style.ToString();
            }
        }


        /// <summary>
        /// 鼠标单击按下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SubTitleFontButtonLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }
        /// <summary>
        /// 鼠标离开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SubTitleFontButtonLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }
        /// <summary>
        /// 鼠标放在上面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SubTitleFontButtonLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }
        /// <summary>
        /// 单位姓名格式选择按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SubTitleFontButtonLabel_Click(object sender, EventArgs e)
        {
            // 创建字体选择对话框
            FontDialog fontDialog = new FontDialog
            {
                AllowScriptChange = false,
                ShowColor = false,
                ShowHelp = false,
                ShowEffects = false
            };

            // 处理用户的字体选择
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取用户所选字体
                Font selectedFont = fontDialog.Font;

                // 在 label1 中显示所选字体的样式
                this.subTitleFontLabel.Text = selectedFont.Name;
                this.subTitleFontSizeLabel.Text = selectedFont.Size.ToString();
                this.subTitleFontStyleLabel.Text = selectedFont.Style.ToString();
            }
        }


        /// <summary>
        /// 鼠标按下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FirstHeadingFontButtonLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }
        /// <summary>
        /// 鼠标离开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FirstHeadingFontButtonLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }
        /// <summary>
        /// 鼠标放在上面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FirstHeadingFontButtonLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }
        /// <summary>
        /// 一级标题格式选择按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FirstHeadingFontButtonLabel_Click(object sender, EventArgs e)
        {
            // 创建字体选择对话框
            FontDialog fontDialog = new FontDialog
            {
                AllowScriptChange = false,
                ShowColor = false,
                ShowHelp = false,
                ShowEffects = false
            };
            // 处理用户的字体选择
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取用户所选字体
                Font selectedFont = fontDialog.Font;
                this.firstHeadingFontLabel.Text = selectedFont.Name;
                this.firstHeadingFontSizeLabel.Text = selectedFont.Size.ToString();
                this.firstHeadingFontStyleLabel.Text = selectedFont.Style.ToString();
            }
        }


        /// <summary>
        /// 鼠标按下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SecondHeadingFontButtonLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }
        /// <summary>
        /// 鼠标离开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SecondHeadingFontButtonLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }
        /// <summary>
        /// 鼠标放在上面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SecondHeadingFontButtonLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }
        /// <summary>
        /// 二级标题格式选择按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SecondHeadingFontButtonLabel_Click(object sender, EventArgs e)
        {
            // 创建字体选择对话框
            FontDialog fontDialog = new FontDialog
            {
                AllowScriptChange = false,
                ShowColor = false,
                ShowHelp = false,
                ShowEffects = false
            };

            // 处理用户的字体选择
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取用户所选字体
                Font selectedFont = fontDialog.Font;

                // 在 label1 中显示所选字体的样式
                this.secondHeadingFontLabel.Text = selectedFont.Name;
                this.secondHeadingFontSizeLabel.Text = selectedFont.Size.ToString();
                this.secondHeadingFontStyleLabel.Text = selectedFont.Style.ToString();
            }
        }


        #endregion


        private void FirstHeadingMinusLabel_Click(object sender, EventArgs e) { FirstHeadingMinusToolStripMenuItem_Click(sender, e); }

        private void FirstHeadingPlusLabel_Click(object sender, EventArgs e) { FirstHeadingPlusToolStripMenuItem_Click(sender, e); }

        private void QuestionStorageMinusLabel_Click(object sender, EventArgs e) { QuestionStorageMinusToolStripMenuItem_Click(sender, e); }

        private void QuestionStoragePlusLabel_Click(object sender, EventArgs e) { QuestionStoragePlusToolStripMenuItem_Click(sender, e); }

        private void FirstHeadingMinusLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }

        private void FirstHeadingMinusLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }

        private void FirstHeadingMinusLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }

        private void FirstHeadingPlusLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }

        private void FirstHeadingPlusLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }

        private void FirstHeadingPlusLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }

        private void QuestionStorageMinusLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }

        private void QuestionStorageMinusLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }

        private void QuestionStorageMinusLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }

        private void QuestionStoragePlusLabel_MouseDown(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientActiveCaption;//按下
        }

        private void QuestionStoragePlusLabel_MouseLeave(object sender, EventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.ControlLight;//离开
        }

        private void QuestionStoragePlusLabel_MouseMove(object sender, MouseEventArgs e)
        {
            Label labelButton = sender as Label;
            labelButton.BackColor = SystemColors.GradientInactiveCaption;//放在上面
        }
    }
}