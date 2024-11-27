//This file is part of ExamPaper Factory
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ExamPaperFactory
{
    public class ExamPaperTask
    {
        private string taskFilePath;//XML文件绝对路径

        private string titleName;//试卷名称

        private string paperSize;//试卷纸张大小
        private string orientation;//试卷排版方向
        private string textColumns;//试卷分栏数量（1/2）

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
            public string path;//路径
            public int amount;//选题数量
        }

        //大题题目列表
        private List<string> questionName;

        //大题每题分值列表
        private List<float> questionGrade;

        //大题下每小题的空行数
        private List<int> questionNextLineSpace;

        //大题完整列表，除了大题名称，包括了里面的所有东西
        private List<List<QuestionElement>> firstQuestionList;

        public string TaskFilePath { get => taskFilePath; set => taskFilePath = value; }

        public string TitleName { get => titleName; set => titleName = value; }

        public string PaperSize { get => paperSize; set => paperSize = value; }
        public string Orientation { get => orientation; set => orientation = value; }
        public string TextColumns { get => textColumns; set => textColumns = value; }


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
        public List<int> QuestionNextLineSpace { get => questionNextLineSpace; set => questionNextLineSpace = value; }
        public List<List<QuestionElement>> FirstQuestionList { get => firstQuestionList; set => firstQuestionList = value; }


        /// <summary>
        /// 读取XML文件所有内容（已更新）
        /// </summary>
        /// <param name="taskFilepath">传入xml文件绝对路径</param>
        public bool ReadXMLFile(string taskFilePath)
        {
            TaskFilePath = taskFilePath;
            if (File.Exists(TaskFilePath))
            {
                try
                {
                    XDocument xdoc = XDocument.Load(taskFilePath);
                    if (xdoc != null && xdoc.Root.Name == "examPaperTask")//判断xml文件内容是否有效
                    {
                        var node = xdoc.Descendants("examPaperTask");
                        titleName = node.FirstOrDefault().Attribute("title").Value;
                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("paperSize").Any())
                            PaperSize = xdoc.Descendants("paperSize").FirstOrDefault().Value;
                        else return false;
                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("textColumns").Any())
                            TextColumns = xdoc.Descendants("textColumns").FirstOrDefault().Value;
                        else return false;
                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("orientation").Any())
                            Orientation = xdoc.Descendants("orientation").FirstOrDefault().Value;
                        else return false;

                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("paperMargin").Descendants("top").Any())
                            MarginTop = float.Parse(xdoc.Descendants("top").FirstOrDefault().Value);
                        else return false;
                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("paperMargin").Descendants("bottom").Any())
                            MarginBottom = float.Parse(xdoc.Descendants("bottom").FirstOrDefault().Value);
                        else return false;
                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("paperMargin").Descendants("left").Any())
                            MarginLeft = float.Parse(xdoc.Descendants("left").FirstOrDefault().Value);
                        else return false;
                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("paperMargin").Descendants("right").Any())
                            MarginRight = float.Parse(xdoc.Descendants("right").FirstOrDefault().Value);
                        else return false;

                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("mainTitle").Any())
                        {
                            var queryMainTitle = from item in xdoc.Element("examPaperTask").Element("form").Element("mainTitle").Elements()
                                                 select new { TypeName = item.Name, Saying = item.Value };
                            foreach (var item in queryMainTitle)
                            {
                                if (item.TypeName.ToString() == "font") { MainTitleFont = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontStyle") { MainTitleFontStyle = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontSize") { MainTitleFontSize = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "nextLineSpace") { MainTitleNextLineSpace = item.Saying.ToString(); }
                            }
                        }
                        else return false;

                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("subTitle").Any())
                        {
                            var querySubTitle = from item in xdoc.Element("examPaperTask").Element("form").Element("subTitle").Elements()
                                                select new { TypeName = item.Name, Saying = item.Value };
                            foreach (var item in querySubTitle)
                            {
                                if (item.TypeName.ToString() == "font") { SubTitleFont = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontStyle") { SubTitleFontStyle = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontSize") { SubTitleFontSize = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "nextLineSpace") { SubTitleNextLineSpace = item.Saying.ToString(); }
                            }
                        }
                        else return false;

                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("firstHeading").Any())
                        {
                            var queryFirstHeading = from item in xdoc.Element("examPaperTask").Element("form").Element("firstHeading").Elements()
                                                    select new { TypeName = item.Name, Saying = item.Value };
                            foreach (var item in queryFirstHeading)
                            {
                                if (item.TypeName.ToString() == "font") { FirstHeadingTitleFont = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontStyle") { FirstHeadingTitleFontStyle = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontSize") { FirstHeadingTitleFontSize = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "nextLineSpace") { FirstHeadingTitleNextLineSpace = item.Saying.ToString(); }
                            }
                        }
                        else return false;

                        if (xdoc.Descendants("examPaperTask").Descendants("form").Descendants("secondHeading").Any())
                        {
                            var querySecondHeading = from item in xdoc.Element("examPaperTask").Element("form").Element("secondHeading").Elements()
                                                     select new { TypeName = item.Name, Saying = item.Value };
                            foreach (var item in querySecondHeading)
                            {
                                if (item.TypeName.ToString() == "font") { SecondHeadingTitleFont = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontStyle") { SecondHeadingTitleFontStyle = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "fontSize") { SecondHeadingTitleFontSize = item.Saying.ToString(); }
                                if (item.TypeName.ToString() == "nextLineSpace") { SecondHeadingTitleNextLineSpace = item.Saying.ToString(); }
                            }
                        }
                        else return false;

                        if (xdoc.Descendants("examPaperTask").Descendants("content").Descendants("question").Any())
                        {
                            IEnumerable<XElement> questionXElements = xdoc.Element("examPaperTask").Element("content").Elements("question");
                            List<XElement> questions = questionXElements.ToList();
                            //Console.WriteLine(questions.ToString());
                            //Console.WriteLine(questions.Count);
                            FirstQuestionList = new List<List<QuestionElement>>();
                            QuestionName = new List<string>();
                            QuestionGrade = new List<float>();
                            QuestionNextLineSpace = new List<int>();
                            foreach (XElement questionItem in questions)//所有的question合集
                            {
                                XAttribute attrQuestionName = questionItem.Attribute("name");
                                QuestionName.Add(attrQuestionName.Value);
                                XAttribute attrQuestionGrade = questionItem.Attribute("grade");
                                QuestionGrade.Add(float.Parse(attrQuestionGrade.Value));
                                XAttribute attrQuestionNextLineSpace = questionItem.Attribute("nextLineSpace");
                                QuestionNextLineSpace.Add(int.Parse(attrQuestionNextLineSpace.Value));

                                IEnumerable<XElement> dirXElements = questionItem.Elements("dir");
                                List<XElement> dirs = dirXElements.ToList();

                                List<QuestionElement> qeList = new List<QuestionElement>();
                                foreach (XElement dirItem in dirs)//所有的dir合集
                                {
                                    QuestionElement qe = new QuestionElement
                                    {
                                        path = dirItem.Attribute("path").Value,
                                        amount = int.Parse(dirItem.Element("amount").Value),
                                    };
                                    //Console.WriteLine("path:{0},grade:{1},amount:{2}", qe.path, qe.grade, qe.amount);
                                    qeList.Add(qe);
                                }
                                FirstQuestionList.Add(qeList);
                            }
                        }
                        else return false;

                        //检测时使用代码，运行成功后不再需要
                        //for (int i = 0; i < FirstQuestionList.Count; i++)
                        //{
                        //    Console.WriteLine("第{0}题:{1}", 1 + i, QuestionName[i]);
                        //    for (int j = 0; j < FirstQuestionList[i].Count; j++)
                        //    {
                        //        string path = FirstQuestionList[i][j].path;
                        //        int amount = FirstQuestionList[i][j].amount;
                        //        int nextLineSpacing = FirstQuestionList[i][j].nextLineSpacing;
                        //        Console.WriteLine("question[{0}][{1}]: path:{2}, grade:{3}, amount:{4}", i, j, path, nextLineSpacing, amount);
                        //    }
                        //}
                        //Console.WriteLine("Title:{0}", this.titleName);
                        return true;
                    }
                    else return false;
                }
                catch { return false; }
            }
            else return false;
        }

        /// <summary>
        /// 将所有涉及到的属性按标准格式写入XML文件（已更新）
        /// </summary>
        /// <param name="taskFilePath">文件保存地址</param>
        /// <returns>返回写入成功情况</returns>
        public bool WriteToXMLFile(string taskFilePath)
        {
            //获取根节点对象
            XDocument document = new XDocument();
            XElement examPaperTaskRoot = new XElement("examPaperTask");//添加根节点
            XAttribute attribute = new XAttribute("title", TitleName);//根节点属性
            examPaperTaskRoot.Add(attribute);//根节点添加属性

            //form节点信息
            XElement formElement = new XElement("form",
                                        new XElement("paperSize", PaperSize),
                                        new XElement("textColumns", TextColumns),
                                        new XElement("orientation", Orientation),
                                        new XElement("paperMargin",
                                            new XElement("top", MarginTop),
                                            new XElement("bottom", MarginBottom),
                                            new XElement("left", MarginLeft),
                                            new XElement("right", MarginRight)),
                                        new XElement("mainTitle",
                                            new XElement("font", MainTitleFont),
                                            new XElement("fontStyle", MainTitleFontStyle),
                                            new XElement("fontSize", MainTitleFontSize),
                                            new XElement("nextLineSpace", MainTitleNextLineSpace)),
                                        new XElement("subTitle",
                                            new XElement("font", SubTitleFont),
                                            new XElement("fontStyle", SubTitleFontStyle),
                                            new XElement("fontSize", SubTitleFontSize),
                                            new XElement("nextLineSpace", SubTitleNextLineSpace)),
                                        new XElement("firstHeading",
                                            new XElement("font", FirstHeadingTitleFont),
                                            new XElement("fontStyle", FirstHeadingTitleFontStyle),
                                            new XElement("fontSize", FirstHeadingTitleFontSize),
                                            new XElement("nextLineSpace", FirstHeadingTitleNextLineSpace)),
                                        new XElement("secondHeading",
                                            new XElement("font", SecondHeadingTitleFont),
                                            new XElement("fontStyle", SecondHeadingTitleFontStyle),
                                            new XElement("fontSize", SecondHeadingTitleFontSize),
                                            new XElement("nextLineSpace", SecondHeadingTitleNextLineSpace)));
            examPaperTaskRoot.Add(formElement);

            //content节点信息
            XElement contentElement = new XElement("content");

            for (int i = 0; i < QuestionName.Count; i++)
            {
                XElement questionElement = new XElement("question");//新建一个question节点
                XAttribute questionNameXAttribute = new XAttribute("name", QuestionName[i]);//question节点属性
                questionElement.Add(questionNameXAttribute);//将属性questionName添加进question节点
                XAttribute questionGradeXAttribute = new XAttribute("grade", QuestionGrade[i]);//question节点属性
                questionElement.Add(questionGradeXAttribute);//将属性questionGrade添加进question节点
                XAttribute questionNextLineSpaceXattribute = new XAttribute("nextLineSpace", QuestionNextLineSpace[i]);//question节点属性
                questionElement.Add(questionNextLineSpaceXattribute);//将属性questionNextLineSpace添加进question节点

                //Console.WriteLine("QuestionName[{0}] = {1}", i, QuestionName[i]);
                for (int j = 0; j < FirstQuestionList[i].Count; j++)
                {
                    //Console.WriteLine("FirstQuestionList[{0}][{1}].path = {2}", i,j, FirstQuestionList[i][j].path);
                    //Console.WriteLine("FirstQuestionList[{0}][{1}].amount = {2}", i,j, FirstQuestionList[i][j].amount);
                    //Console.WriteLine("FirstQuestionList[{0}][{1}].nextLineSpace = {2}", i, j, FirstQuestionList[i][j].nextLineSpace);
                    XElement dirElement = new XElement("dir",
                                            new XElement("amount", FirstQuestionList[i][j].amount));
                    XAttribute dirXAttribute = new XAttribute("path", FirstQuestionList[i][j].path);
                    dirElement.Add(dirXAttribute);
                    questionElement.Add(dirElement);
                }
                contentElement.Add(questionElement);
            }

            examPaperTaskRoot.Add(contentElement);//添加content节点
            examPaperTaskRoot.Save(taskFilePath);
            return true;
        }

        /// <summary>
        /// 读取新建文件后所输入的所有值（已更新）
        /// </summary>
        internal void ReadNewExamTaskContentFormData(NewExamTaskContentForm newExamTaskContentForm)
        {
            taskFilePath = null;//定义为空
            TitleName = newExamTaskContentForm.TitleName;

            PaperSize = newExamTaskContentForm.PaperSize;
            TextColumns = newExamTaskContentForm.TextColumns;
            Orientation = newExamTaskContentForm.Orientation;

            MarginTop = newExamTaskContentForm.MarginTop;
            MarginBottom = newExamTaskContentForm.MarginBottom;
            MarginLeft = newExamTaskContentForm.MarginLeft;
            MarginRight = newExamTaskContentForm.MarginRight;

            MainTitleFont = newExamTaskContentForm.MainTitleFont;
            MainTitleFontSize = newExamTaskContentForm.MainTitleFontSize;
            MainTitleFontStyle = newExamTaskContentForm.MainTitleFontStyle;
            MainTitleNextLineSpace = newExamTaskContentForm.MainTitleNextLineSpace;

            SubTitleFont = newExamTaskContentForm.SubTitleFont;
            SubTitleFontSize = newExamTaskContentForm.SubTitleFontSize;
            SubTitleFontStyle = newExamTaskContentForm.SubTitleFontStyle;
            SubTitleNextLineSpace = newExamTaskContentForm.SubTitleNextLineSpace;

            SecondHeadingTitleFont = newExamTaskContentForm.SecondHeadingTitleFont;
            SecondHeadingTitleFontSize = newExamTaskContentForm.SecondHeadingTitleFontSize;
            SecondHeadingTitleFontStyle = newExamTaskContentForm.SecondHeadingTitleFontStyle;
            SecondHeadingTitleNextLineSpace = newExamTaskContentForm.SecondHeadingTitleNextLineSpace;

            FirstHeadingTitleFont = newExamTaskContentForm.FirstHeadingTitleFont;
            FirstHeadingTitleFontSize = newExamTaskContentForm.FirstHeadingTitleFontSize;
            FirstHeadingTitleFontStyle = newExamTaskContentForm.FirstHeadingTitleFontStyle;
            FirstHeadingTitleNextLineSpace = newExamTaskContentForm.FirstHeadingTitleNextLineSpace;

            QuestionName = newExamTaskContentForm.QuestionName;
            QuestionGrade = newExamTaskContentForm.QuestionGrade;
            QuestionNextLineSpace = newExamTaskContentForm.QuestionNextLineSpace;

            FirstQuestionList = new List<List<QuestionElement>>();
            for (int i = 0; i < newExamTaskContentForm.FirstQuestionList.Count; i++)
            {
                List<QuestionElement> qeList = new List<QuestionElement>();
                for (int j = 0; j < newExamTaskContentForm.FirstQuestionList[i].Count; j++)
                {
                    QuestionElement qe = new QuestionElement();
                    qe.path = newExamTaskContentForm.FirstQuestionList[i][j].Path;
                    qe.amount = newExamTaskContentForm.FirstQuestionList[i][j].Amount;
                    qeList.Add(qe);
                }
                FirstQuestionList.Add(qeList);
            }

            //for (int i = 0; i < FirstQuestionList.Count; i++)
            //{
            //    Console.WriteLine("第{0}题:{1}", 1 + i, QuestionName[i]);
            //    for (int j = 0; j < FirstQuestionList[i].Count; j++)
            //    {
            //        string path = FirstQuestionList[i][j].path;
            //        int amount = FirstQuestionList[i][j].amount;
            //        Console.WriteLine("question[{0}][{1}]: path:{2}, amount:{3}", i, j, path, amount);
            //    }
            //}
        }

        /// <summary>
        /// 读取新建文件后所输入的所有值（已更新）
        /// </summary>
        internal void ReadEditExamTaskContentFormData(EditExamTaskContentForm editExamTaskContentForm)
        {
            taskFilePath = null;//定义为空
            TitleName = editExamTaskContentForm.TitleName;

            PaperSize = editExamTaskContentForm.PaperSize;
            textColumns = editExamTaskContentForm.TextColumns;
            Orientation = editExamTaskContentForm.Orientation;

            MarginTop = editExamTaskContentForm.MarginTop;
            MarginBottom = editExamTaskContentForm.MarginBottom;
            MarginLeft = editExamTaskContentForm.MarginLeft;
            MarginRight = editExamTaskContentForm.MarginRight;

            MainTitleFont = editExamTaskContentForm.MainTitleFont;
            MainTitleFontSize = editExamTaskContentForm.MainTitleFontSize;
            MainTitleFontStyle = editExamTaskContentForm.MainTitleFontStyle;
            MainTitleNextLineSpace = editExamTaskContentForm.MainTitleNextLineSpace;

            SubTitleFont = editExamTaskContentForm.SubTitleFont;
            SubTitleFontSize = editExamTaskContentForm.SubTitleFontSize;
            SubTitleFontStyle = editExamTaskContentForm.SubTitleFontStyle;
            SubTitleNextLineSpace = editExamTaskContentForm.SubTitleNextLineSpace;

            SecondHeadingTitleFont = editExamTaskContentForm.SecondHeadingTitleFont;
            SecondHeadingTitleFontSize = editExamTaskContentForm.SecondHeadingTitleFontSize;
            SecondHeadingTitleFontStyle = editExamTaskContentForm.SecondHeadingTitleFontStyle;
            SecondHeadingTitleNextLineSpace = editExamTaskContentForm.SecondHeadingTitleNextLineSpace;

            FirstHeadingTitleFont = editExamTaskContentForm.FirstHeadingTitleFont;
            FirstHeadingTitleFontSize = editExamTaskContentForm.FirstHeadingTitleFontSize;
            FirstHeadingTitleFontStyle = editExamTaskContentForm.FirstHeadingTitleFontStyle;
            FirstHeadingTitleNextLineSpace = editExamTaskContentForm.FirstHeadingTitleNextLineSpace;

            QuestionName = editExamTaskContentForm.QuestionName;
            QuestionGrade = editExamTaskContentForm.QuestionGrade;
            QuestionNextLineSpace = editExamTaskContentForm.QuestionNextLineSpace;

            FirstQuestionList = new List<List<QuestionElement>>();
            for (int i = 0; i < editExamTaskContentForm.FirstQuestionList.Count; i++)
            {
                List<QuestionElement> qeList = new List<QuestionElement>();
                for (int j = 0; j < editExamTaskContentForm.FirstQuestionList[i].Count; j++)
                {
                    QuestionElement qe = new QuestionElement();
                    qe.path = editExamTaskContentForm.FirstQuestionList[i][j].Path;
                    qe.amount = editExamTaskContentForm.FirstQuestionList[i][j].Amount;
                    qeList.Add(qe);
                }
                FirstQuestionList.Add(qeList);
            }

            //for (int i = 0; i < FirstQuestionList.Count; i++)
            //{
            //    Console.WriteLine("第{0}题:{1}", 1 + i, QuestionName[i]);
            //    for (int j = 0; j < FirstQuestionList[i].Count; j++)
            //    {
            //        string path = FirstQuestionList[i][j].path;
            //        double amount = FirstQuestionList[i][j].amount;
            //        double nextLineSpace = FirstQuestionList[i][j].nextLineSpace;
            //        Console.WriteLine("question[{0}][{1}]: path:{2}, amount:{3}, nextLineSpace:{4}", i, j, path, amount, nextLineSpace);
            //    }
            //}
        }
    }
}