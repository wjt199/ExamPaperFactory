//This file is part of ExamPaper Factory
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
namespace ExamPaperFactory
{
    /// <summary>
    /// 用来生成试卷的类
    /// </summary>
    internal class ExamPaper
    {
        private string examPaperFileFormTime;
        private ExamPaperTask examPaperTask;//生成试卷的任务
        private string examPaperDirPath = null;//生成试卷的目录位置
        private string examPaperName = null;//生成试卷的文件名称
        private string examPaperAnswerName = null;//生成试卷答案的文件名称
        private string mainFormTitle = null;//试卷任务标题

        public ExamPaperTask ExamPaperTask1 { get => examPaperTask; set => examPaperTask = value; }
        public string ExamPaperDirPath { get => examPaperDirPath; set => examPaperDirPath = value; }
        public string ExamPaperName { get => examPaperName; set => examPaperName = value; }
        public string ExamPaperFileFormTime { get => examPaperFileFormTime; set => examPaperFileFormTime = value; }
        public string ExamPaperAnswerName { get => examPaperAnswerName; set => examPaperAnswerName = value; }
        public string MainFormTitle { get => mainFormTitle; set => mainFormTitle = value; }

        /// <summary>
        /// 检查数据的合法性
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        internal bool CheckIsValid()
        {
            return true;
        }

        /// <summary>
        /// 生成试卷
        /// </summary>
        /// <returns>生成成功返回true，失败返回false</returns>
        internal bool FormExamPaper(BackgroundWorker backgroundWorker)
        {
            backgroundWorker.ReportProgress(5);
            //检查并创建生成试卷和答案的目录，反之非法行为导致生成失败
            if (!Directory.Exists(examPaperDirPath)) Directory.CreateDirectory(examPaperDirPath);
            ExamPaperName = MainFormTitle + "（试卷）" + ExamPaperFileFormTime + ".docx";
            //Console.WriteLine("ExamPaperName:{0}", ExamPaperName);
            //生成试卷
            //Word应用程序变量
            Application wordexamop = null;
            Document wordexam = null;
            backgroundWorker.ReportProgress(10);
            try
            {
                wordexamop = new Application();
                wordexamop.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;//屏蔽宏
                wordexamop.Visible = false;//使文档不可见,生成文件过程中文档不可见
                                           //wordexamop.Visible = true;//使文档可见

                Object missing = Missing.Value;//由于使用的是COM库，因此有许多变量需要使用Missing.Value代替

                wordexam = wordexamop.Documents.Add();
                backgroundWorker.ReportProgress(20);
                //纸张大小设置
                if (examPaperTask.PaperSize == "A3") wordexam.PageSetup.PaperSize = WdPaperSize.wdPaperA3;//设置纸张样式为A3纸
                else if (examPaperTask.PaperSize == "A4") wordexam.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置纸张样式为A4纸
                else wordexam.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//默认纸张样式为A4规格

                //横竖版面设置
                if (examPaperTask.Orientation == "竖版") wordexam.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                else if (examPaperTask.Orientation == "横版") wordexam.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                else wordexam.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

                //页面设置
                wordexam.PageSetup.LeftMargin = float.Parse(examPaperTask.MarginLeft.ToString()); //2.82CM
                wordexam.PageSetup.RightMargin = float.Parse(examPaperTask.MarginRight.ToString());
                wordexam.PageSetup.TopMargin = float.Parse(examPaperTask.MarginTop.ToString());//3.525CM
                wordexam.PageSetup.BottomMargin = float.Parse(examPaperTask.MarginBottom.ToString());

                //分栏设置
                if (examPaperTask.TextColumns == "2") wordexam.PageSetup.TextColumns.SetCount(2);

                //添加试卷标题
                wordexamop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordexamop.Selection.Font.Size = float.Parse(examPaperTask.MainTitleFontSize);
                wordexamop.Selection.Font.Name = examPaperTask.MainTitleFont;
                wordexamop.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;//首行缩进为0
                wordexamop.Selection.ParagraphFormat.FirstLineIndent = wordexam.Application.CentimetersToPoints((float)0.75 * 0);//首行缩进为0，全局设置为0
                wordexamop.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                wordexamop.Selection.ParagraphFormat.LineSpacing = float.Parse(examPaperTask.MainTitleNextLineSpace);
                wordexam.Paragraphs.Last.Range.Text = this.examPaperTask.TitleName + Environment.NewLine + Environment.NewLine;//试卷标题

                object unite = WdUnits.wdStory;
                wordexamop.Selection.EndKey(ref unite, ref missing);//将光标移动到文本末尾
                wordexamop.Selection.Font.Size = float.Parse(examPaperTask.SubTitleFontSize);//次标题大小
                wordexamop.Selection.Font.Name = examPaperTask.SubTitleFont;
                wordexam.Paragraphs.Last.Range.Text = "单位：___________    姓名：___________    成绩：___________" + Environment.NewLine + Environment.NewLine;
                backgroundWorker.ReportProgress(40);
                int step = 40 / examPaperTask.QuestionName.Count;
                for (int i = 0; i < examPaperTask.QuestionName.Count; i++)
                {
                    //题目标题加内容
                    wordexamop.Selection.EndKey(ref unite, ref missing);//将光标移动到文本末尾
                                                                        //wordexamop.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordexamop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    wordexamop.Selection.Font.Size = float.Parse(examPaperTask.FirstHeadingTitleFontSize);//四号字体大小
                    wordexamop.Selection.Font.Name = examPaperTask.FirstHeadingTitleFont;
                    //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值,前面已经设置过了
                    wordexamop.Selection.ParagraphFormat.LineSpacing = float.Parse(examPaperTask.FirstHeadingTitleNextLineSpace);

                    float sum = 0;
                    for (int j = 0; j < examPaperTask.FirstQuestionList[i].Count; j++) sum += examPaperTask.QuestionGrade[i] * examPaperTask.FirstQuestionList[i][j].amount;

                    wordexam.Paragraphs.Last.Range.Text = Utilities.IntToBigNum(i + 1) + examPaperTask.QuestionName[i] + "（每题" + examPaperTask.QuestionGrade[i] + "分，共 " + sum + "分）" + Environment.NewLine;
                    //i++;
                    wordexamop.Selection.EndKey(ref unite, ref missing);//将光标移动到文本末尾
                    wordexamop.Selection.Font.Size = float.Parse(examPaperTask.SecondHeadingTitleFontSize);//四号字体大小
                    wordexamop.Selection.Font.Name = examPaperTask.SecondHeadingTitleFont;
                    //wordexamop.Selection.Font.;
                    //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                    wordexamop.Selection.ParagraphFormat.LineSpacing = float.Parse(examPaperTask.SecondHeadingTitleNextLineSpace);//固定行距为18

                    //Console.WriteLine("examPaperTask.FirstQuestionList[{0}].Count = {1}", i, examPaperTask.FirstQuestionList[i].Count);
                    for (int j = 0; j < examPaperTask.FirstQuestionList[i].Count; j++)
                    {
                        List<List<string>> questionList = new List<List<string>>();
                        questionList = null;
                        while (j < examPaperTask.FirstQuestionList[i].Count)
                        {
                            List<List<string>> newQuestionList = new List<List<string>>();
                            string questionStoragePath = examPaperTask.FirstQuestionList[i][j].path;
                            int num = examPaperTask.FirstQuestionList[i][j].amount;
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
                            for (int m = 0; m < examPaperTask.QuestionNextLineSpace[i]; m++)//每个大的题型的小题行距以每个大题第一个题库输入的行间距为准，后续题库的输入值不予采纳
                                wordexam.Paragraphs.Last.Range.Text = Environment.NewLine;
                        }
                    }
                    wordexam.Paragraphs.Last.Range.Text = Environment.NewLine;//大题之间的空行
                    backgroundWorker.ReportProgress(40 + step);
                }
                backgroundWorker.ReportProgress(80);

                //Console.WriteLine(Path.Combine(ExamPaperDirPath, ExamPaperName));
                wordexam.SaveAs(Path.Combine(ExamPaperDirPath, ExamPaperName));
                wordexam.Close(ref missing, ref missing, ref missing);
                wordexam = null;

                wordexamop.Quit(ref missing, ref missing, ref missing);
                wordexamop = null;
                backgroundWorker.ReportProgress(90);
                Thread.Sleep(500);
                backgroundWorker.ReportProgress(99);
                Thread.Sleep(500);
                backgroundWorker.ReportProgress(100);
                return true;
            }
            catch
            {
                //Console.WriteLine("发生了异常");
                return false;
            }
        }

        /// <summary>
        /// 生成试卷+答案
        /// </summary>
        /// <returns>生成成功返回true，失败返回false</returns>
        internal bool FormExamPaperAndAnswer(BackgroundWorker backgroundWorker)
        {
            backgroundWorker.ReportProgress(1);

            List<List<int[]>> chosenIndex = new List<List<int[]>>();//用来装试卷和答案序号的列表
            backgroundWorker.ReportProgress(5);
            //检查并创建生成试卷和答案的目录，反之非法行为导致生成失败
            if (!Directory.Exists(examPaperDirPath)) Directory.CreateDirectory(examPaperDirPath);

            ///生成试卷
            ExamPaperName = MainFormTitle + "（试卷）" + ExamPaperFileFormTime + ".docx";
            //Console.WriteLine("ExamPaperName:{0}", ExamPaperName);

            //Word应用程序变量
            Application wordexampaperop = null;
            Document wordexampaper = null;
            backgroundWorker.ReportProgress(8);
            try
            {
                wordexampaperop = new Application();
                wordexampaperop.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;//屏蔽宏
                wordexampaperop.Visible = false;//使文档不可见,生成文件过程中文档不可见
                                                //wordexamop.Visible = true;//使文档可见

                Object exampapermissing = Missing.Value;//由于使用的是COM库，因此有许多变量需要使用Missing.Value代替

                wordexampaper = wordexampaperop.Documents.Add();

                //纸张大小设置
                if (examPaperTask.PaperSize == "A3") wordexampaper.PageSetup.PaperSize = WdPaperSize.wdPaperA3;//设置纸张样式为A3纸
                else if (examPaperTask.PaperSize == "A4") wordexampaper.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置纸张样式为A4纸
                else wordexampaper.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//默认纸张样式为A4规格

                //横竖版面设置
                if (examPaperTask.Orientation == "竖版") wordexampaper.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                else if (examPaperTask.Orientation == "横版") wordexampaper.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                else wordexampaper.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

                //页面设置
                wordexampaper.PageSetup.LeftMargin = float.Parse(examPaperTask.MarginLeft.ToString()); //2.82CM
                wordexampaper.PageSetup.RightMargin = float.Parse(examPaperTask.MarginRight.ToString());
                wordexampaper.PageSetup.TopMargin = float.Parse(examPaperTask.MarginTop.ToString());//3.525CM
                wordexampaper.PageSetup.BottomMargin = float.Parse(examPaperTask.MarginBottom.ToString());

                //分栏设置
                if (examPaperTask.TextColumns == "2") wordexampaper.PageSetup.TextColumns.SetCount(2);
                else wordexampaper.PageSetup.TextColumns.SetCount(1);

                //添加试卷标题
                wordexampaperop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordexampaperop.Selection.Font.Size = float.Parse(examPaperTask.MainTitleFontSize);
                wordexampaperop.Selection.Font.Name = examPaperTask.MainTitleFont;
                wordexampaperop.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;//首行缩进为0
                wordexampaperop.Selection.ParagraphFormat.FirstLineIndent = wordexampaper.Application.CentimetersToPoints((float)0.75 * 0);//首行缩进为0，全局设置为0
                wordexampaperop.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                wordexampaperop.Selection.ParagraphFormat.LineSpacing = float.Parse(examPaperTask.MainTitleNextLineSpace);
                wordexampaper.Paragraphs.Last.Range.Text = this.examPaperTask.TitleName + Environment.NewLine + Environment.NewLine;//试卷标题
                backgroundWorker.ReportProgress(10);
                object exampaperunite = WdUnits.wdStory;
                wordexampaperop.Selection.EndKey(ref exampaperunite, ref exampapermissing);//将光标移动到文本末尾
                wordexampaperop.Selection.Font.Size = float.Parse(examPaperTask.SubTitleFontSize);//次标题大小
                wordexampaperop.Selection.Font.Name = examPaperTask.SubTitleFont;
                wordexampaper.Paragraphs.Last.Range.Text = "单位：___________    姓名：___________    成绩：___________" + Environment.NewLine + Environment.NewLine;
                backgroundWorker.ReportProgress(20);
                int mount1 = examPaperTask.QuestionName.Count;
                int step1 = 30 / mount1;
                for (int i = 0; i < examPaperTask.QuestionName.Count; i++)
                {
                    //题目标题加内容
                    wordexampaperop.Selection.EndKey(ref exampaperunite, ref exampapermissing);//将光标移动到文本末尾
                    wordexampaperop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    wordexampaperop.Selection.Font.Size = float.Parse(examPaperTask.FirstHeadingTitleFontSize);//大题字体大小设置
                    wordexampaperop.Selection.Font.Name = examPaperTask.FirstHeadingTitleFont;//大题字体设置
                                                                                              //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值,前面已经设置过了
                    wordexampaperop.Selection.ParagraphFormat.LineSpacing = float.Parse(examPaperTask.FirstHeadingTitleNextLineSpace);//大题行间距设置
                    float sum = 0;
                    for (int j = 0; j < examPaperTask.FirstQuestionList[i].Count; j++) sum += examPaperTask.QuestionGrade[i] * examPaperTask.FirstQuestionList[i][j].amount;
                    wordexampaper.Paragraphs.Last.Range.Text = Utilities.IntToBigNum(i + 1) + examPaperTask.QuestionName[i] + "（每题" + examPaperTask.QuestionGrade[i] + "分，共" + sum + "分）" + Environment.NewLine;
                    wordexampaperop.Selection.EndKey(ref exampaperunite, ref exampapermissing);//将光标移动到文本末尾
                    wordexampaperop.Selection.Font.Size = float.Parse(examPaperTask.SecondHeadingTitleFontSize);//小题字体大小设置
                    wordexampaperop.Selection.Font.Name = examPaperTask.SecondHeadingTitleFont;//小题字体设置
                                                                                               //wordexamop.Selection.Font.;
                                                                                               //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                    wordexampaperop.Selection.ParagraphFormat.LineSpacing = float.Parse(examPaperTask.SecondHeadingTitleNextLineSpace);//小题行间距设置

                    for (int j = 0; j < examPaperTask.FirstQuestionList[i].Count; j++)
                    {
                        List<List<string>> questionList = new List<List<string>>();
                        //questionList = null;
                        List<int[]> ints = new List<int[]>();
                        while (j < examPaperTask.FirstQuestionList[i].Count)
                        {
                            List<List<string>> newQuestionList = new List<List<string>>();
                            string questionStoragePath = examPaperTask.FirstQuestionList[i][j].path;//题库路径
                            int num = examPaperTask.FirstQuestionList[i][j].amount;//上述路径对应的选题数
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
                            for (int m = 0; m < examPaperTask.QuestionNextLineSpace[i]; m++)//每个大的题型的小题行距以每个大题第一个题库输入的行间距为准，后续题库的输入值不予采纳
                                wordexampaper.Paragraphs.Last.Range.Text = Environment.NewLine;
                        }

                    }
                    wordexampaper.Paragraphs.Last.Range.Text = Environment.NewLine;//大题之间的空行
                    backgroundWorker.ReportProgress(20 + i * step1);
                }
                wordexampaper.SaveAs(Path.Combine(ExamPaperDirPath, ExamPaperName));//存储试卷
                wordexampaper.Close(ref exampapermissing, ref exampapermissing, ref exampapermissing);
                wordexampaper = null;
                wordexampaperop.Quit(ref exampapermissing, ref exampapermissing, ref exampapermissing);
                wordexampaperop = null;

                backgroundWorker.ReportProgress(50);


                ///生成答案
                ExamPaperAnswerName = MainFormTitle + "（答案）" + ExamPaperFileFormTime + ".docx";
                //Console.WriteLine("ExamPaperName:{0}", ExamPaperName);

                //Word应用程序变量
                Application wordexampaperanswerop = null;
                Document wordexampaperanser = null;

                wordexampaperanswerop = new Application();
                wordexampaperanswerop.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;//屏蔽宏
                wordexampaperanswerop.Visible = false;//使文档不可见,生成文件过程中文档不可见
                                                      //wordexamop.Visible = true;//使文档可见
                backgroundWorker.ReportProgress(55);
                Object exampaperanswermissing = Missing.Value;//由于使用的是COM库，因此有许多变量需要使用Missing.Value代替

                wordexampaperanser = wordexampaperanswerop.Documents.Add();

                //纸张大小设置
                wordexampaperanser.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置纸张样式为A4纸
                //横竖版面设置
                wordexampaperanser.PageSetup.Orientation = WdOrientation.wdOrientPortrait;//设置为竖版
                //页面设置
                wordexampaperanser.PageSetup.LeftMargin = float.Parse(examPaperTask.MarginLeft.ToString()); //2.82CM
                wordexampaperanser.PageSetup.RightMargin = float.Parse(examPaperTask.MarginRight.ToString());
                wordexampaperanser.PageSetup.TopMargin = float.Parse(examPaperTask.MarginTop.ToString());//3.525CM
                wordexampaperanser.PageSetup.BottomMargin = float.Parse(examPaperTask.MarginBottom.ToString());
                backgroundWorker.ReportProgress(60);
                //添加试卷标题
                wordexampaperanswerop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordexampaperanswerop.Selection.Font.Size = float.Parse(examPaperTask.MainTitleFontSize);
                wordexampaperanswerop.Selection.Font.Name = examPaperTask.MainTitleFont;
                wordexampaperanswerop.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0;//首行缩进为0
                wordexampaperanswerop.Selection.ParagraphFormat.FirstLineIndent = wordexampaperanser.Application.CentimetersToPoints((float)0.75 * 0);//首行缩进为0，全局设置为0
                wordexampaperanswerop.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;//行距为固定值
                wordexampaperanswerop.Selection.ParagraphFormat.LineSpacing = float.Parse(examPaperTask.MainTitleNextLineSpace);
                wordexampaperanser.Paragraphs.Last.Range.Text = this.examPaperTask.TitleName + "答案" + Environment.NewLine + Environment.NewLine;//试卷标题
                backgroundWorker.ReportProgress(70);
                object exampaperanswerunite = WdUnits.wdStory;
                wordexampaperanswerop.Selection.EndKey(ref exampaperanswerunite, ref exampaperanswermissing);//将光标移动到文本末尾

                int step2 = 25 / mount1;

                for (int i = 0; i < examPaperTask.QuestionName.Count; i++)
                {
                    //题目标题加内容
                    wordexampaperanswerop.Selection.EndKey(ref exampaperanswerunite, ref exampaperanswermissing);
                    wordexampaperanswerop.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    wordexampaperanswerop.Selection.Font.Size = 10;
                    wordexampaperanswerop.Selection.Font.Name = "黑体";
                    //wordexamop.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;//行距为固定值,前面已经设置过了
                    wordexampaperanswerop.Selection.ParagraphFormat.LineSpacing = 14f;
                    wordexampaperanser.Paragraphs.Last.Range.Text = Utilities.IntToBigNum(i + 1) + examPaperTask.QuestionName[i] + Environment.NewLine;
                    wordexampaperanswerop.Selection.EndKey(ref exampaperanswerunite, ref exampaperanswermissing);//将光标移动到文本末尾
                    wordexampaperanswerop.Selection.Font.Size = 9;
                    wordexampaperanswerop.Selection.Font.Name = "宋体";//小答案内容字体设置
                    wordexampaperanswerop.Selection.ParagraphFormat.LineSpacing = 12f;//小题行间距设置

                    int step = 30 / examPaperTask.FirstQuestionList[i].Count;
                    for (int j = 0; j < examPaperTask.FirstQuestionList[i].Count; j++)
                    {
                        List<List<string>> answerList = new List<List<string>>();
                        answerList = null;
                        while (j < examPaperTask.FirstQuestionList[i].Count)
                        {
                            List<List<string>> newAnswerList = new List<List<string>>();
                            string questionStoragePath = examPaperTask.FirstQuestionList[i][j].path;
                            int num = examPaperTask.FirstQuestionList[i][j].amount;
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
                    backgroundWorker.ReportProgress(70 + i * step2);
                }
                backgroundWorker.ReportProgress(95);

                wordexampaperanser.SaveAs(Path.Combine(ExamPaperDirPath, ExamPaperAnswerName));//存储试卷答案
                wordexampaperanser.Close(ref exampaperanswermissing, ref exampaperanswermissing, ref exampaperanswermissing);
                wordexampaperanser = null;
                wordexampaperanswerop.Quit(ref exampaperanswermissing, ref exampaperanswermissing, ref exampaperanswermissing);
                wordexampaperanswerop = null;

                Thread.Sleep(500);
                backgroundWorker.ReportProgress(99);
                Thread.Sleep(500);
                backgroundWorker.ReportProgress(100);
                return true;
            }
            catch
            {
                Console.WriteLine("发生了异常");
                return false;
            }
        }
    }
}
