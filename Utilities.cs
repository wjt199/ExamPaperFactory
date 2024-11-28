//This file is part of ExamPaper Factory
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;


namespace ExamPaperFactory
{
    internal static class Utilities

    {
        /// <summary>
        /// 根据输入的文件路径类型对ConfigPath.txt文件进行更新
        /// 当文件存在时，进行更新操作，文件不存在时，新建ConfigPath.txt文件并写入新条目
        /// </summary>
        /// <param name="type">文件路径类型(不带“：”)</param>
        /// <param name="dir">路径（绝对路径）</param>
        public static void UpdateConfigPath(string type, string dir)
        {
            List<string> newAllConfigPath = new List<string>();
            if (File.Exists(@"ConfigPath.txt"))
            {
                string[] allConfigPath = File.ReadAllLines(@"ConfigPath.txt");
                if (allConfigPath.Length != 0)
                {
                    foreach (string item in allConfigPath)
                    {
                        if (!item.Contains(type)) { newAllConfigPath.Add(item); }
                    }
                }
                newAllConfigPath.Add(type + ":" + dir);
            }
            else { newAllConfigPath.Add(type + ":" + dir); }
            File.WriteAllLines(@"ConfigPath.txt", newAllConfigPath.ToArray<string>(), Encoding.UTF8);
        }

        /// <summary>
        /// 将数字转换为从1到该数字组成的字符串数组
        /// </summary>
        /// <param name="num">大于1的整数</param>
        /// <returns>字符串数组</returns>
        public static String[] NumToStringList(int num)
        {
            List<string> strList = new List<String>();
            for (int i = 1; i <= num; i++) { strList.Add(i.ToString()); }
            return strList.ToArray();
        }

        /// <summary>
        /// 将两个数字之间的所有整数转换为由小到大的字符串数组
        /// </summary>
        /// <param name="startNum">起始数字</param>
        /// <param name="endNum">结束数字</param>
        /// <returns>字符串数组</returns>
        public static String[] NumToStringList(int startNum, int endNum)
        {
            List<string> strList = new List<String>();
            for (int i = startNum; i <= endNum; i++) { strList.Add(i.ToString()); }
            return strList.ToArray();
        }

        /// <summary>
        /// 将两个数字之间的数字按照给定的步进值转换为由小到大的字符串数组
        /// </summary>
        /// <param name="startNum">起始数字</param>
        /// <param name="endNum">结束数字</param>
        /// <param name="step">步进值</param>
        /// <returns>字符串数组</returns>
        public static String[] NumToStringList(float startNum, float endNum, float step)
        {
            List<string> strList = new List<String>();
            for (float i = startNum; i <= endNum; i += step) { strList.Add(i.ToString()); }
            return strList.ToArray();
        }


        /// <summary>
        /// 读取题库文件，从中获取到题目的数量
        /// </summary>
        /// <param name="questionStoragePath"></param>
        /// <returns></returns>
        public static int CountQuestionNumFromQuestionStoragePath(string questionStoragePath)
        {
            int num = 0;
            StreamReader streamReader = File.OpenText(questionStoragePath);
            string nextLine;
            while ((nextLine = streamReader.ReadLine()) != null)
            {
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) num++;
            }
            streamReader.Close();
            return num;
        }

        /// <summary>
        /// 检查题库是否符合要求,检查方法为：
        /// 检查包含题目和包含答案的行交替进行，而且最开始一定要是“题目”行，最后一定要是“答案”行，
        /// 中间可以有空行，也可以不空行。
        /// 默认为：1.题目行后至最近的答案行前所有内容均为题目内容（包括空行）；
        ///         2.答案行后至最近的题目行前或读取结束前均为答案内容（包括空行）
        /// 满足以上规则，则认为题库文件正确。
        /// </summary>
        /// <param name="questionStoragePath"></param>
        /// <returns></returns>
        internal static bool CheckQuestionStorageFormat(string questionStoragePath)
        {
            //Console.WriteLine("questionStoragePath:{0}", questionStoragePath);
            if (!File.Exists(questionStoragePath)) return false;///检查路径文件存在
            //if (CountQuestionNumFromQuestionStoragePath(questionStoragePath) == 0) return false;
            //return true;
            StreamReader streamReader = File.OpenText(questionStoragePath);//打开文件

            string nextLine;
            int lineIndex = 0;//行号计数器

            int questionStart, questionEnd;
            int answerStart, answerEnd;

            List<List<string>> questionList = new List<List<string>>();
            List<List<string>> answerList = new List<List<string>>();

            List<int> questionIndexList = new List<int>();
            List<int> answerIndexList = new List<int>();

            nextLine = streamReader.ReadLine();//nextline 读取第一行

            lineIndex++;//总行号计数器+1

            while (nextLine != null)
            {
                //string templine = nextLine;
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目："))
                {
                    questionIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine);
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        list.Add(nextLine);
                    }
                    questionList.Add(list);
                }
                else if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案："))
                {
                    answerIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine);
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        //Console.WriteLine("nextLine:{0}", nextLine);
                        //Console.WriteLine("nextLine.Substring(0, 3):{0}", nextLine.Substring(0, 3));
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                        if (nextLine != "") list.Add(nextLine);//答案不要空行
                    }
                    answerList.Add(list);
                }
                else//对于没有出现答案和题目的情况，如下处理
                {
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                    }
                }
            }
            streamReader.Close();

            if (questionIndexList.Count == 0 || answerIndexList.Count == 0) return false;//表明里面没有数据内容，故不符合题库条件
            if (questionIndexList.Count == answerIndexList.Count)
            {
                questionStart = questionIndexList[0];
                answerStart = answerIndexList[0];
                questionEnd = questionIndexList[questionIndexList.Count - 1];
                answerEnd = answerIndexList[answerIndexList.Count - 1];
                if (questionStart > answerStart || questionEnd > answerEnd) return false;
                else return true;
            }
            else return false;
        }

        /// <summary>
        /// 根据题库路径，随机选择给型数量的题目,且题目内容里面不包含“题目：”两个字，同时生成对应的答案。
        /// </summary>
        /// <param name="questionStoragePath">单个题库的路径</param>
        /// <param name="num">该题库选题题数</param>
        /// <returns>Item1:选中的题库内容;Item2:对应的答案内容</returns>
        internal static (List<List<string>>, List<List<string>>) GetNeededQuestionAndAnswerFromQuestionStorage(string questionStoragePath, int num)
        {
            //Console.WriteLine("questionStoragePath:{0}", questionStoragePath);
            if (!File.Exists(questionStoragePath)) return (null, null);///检查路径文件存在
            StreamReader streamReader = File.OpenText(questionStoragePath);//打开文件

            string nextLine;
            int lineIndex = 0;//行号计数器

            List<List<string>> questionList = new List<List<string>>();
            List<List<string>> answerList = new List<List<string>>();

            List<int> questionIndexList = new List<int>();
            List<int> answerIndexList = new List<int>();

            nextLine = streamReader.ReadLine();//nextline 读取第一行

            lineIndex++;//总行号计数器+1

            while (nextLine != null)
            {
                //string templine = nextLine;
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目："))
                {
                    questionIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        list.Add(nextLine);
                    }
                    questionList.Add(list);
                }
                else if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案："))
                {
                    answerIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        //Console.WriteLine("nextLine:{0}", nextLine);
                        //Console.WriteLine("nextLine.Substring(0, 3):{0}", nextLine.Substring(0, 3));
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                        if (nextLine != "") list.Add(nextLine);//答案不要空行
                    }
                    answerList.Add(list);
                }
                else//对于没有出现答案和题目的情况，如下处理
                {
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if ((nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        if ((nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                    }
                }
            }
            streamReader.Close();

            if (questionIndexList.Count == 0 || answerIndexList.Count == 0) return (null, null);//表明里面没有数据内容，故不符合题库条件

            int[] numIndex = new int[num];
            Random random = new Random();
            int num_random;
            int counter = 0;
            while (counter < num)
            {
                num_random = random.Next(1, questionList.Count + 1);
                if (!numIndex.Contains<int>(num_random))
                {
                    numIndex[counter] = num_random;
                    counter++;
                }
            }

            //Console.WriteLine("numIndex.length = {0}", numIndex.Length);
            //Console.WriteLine("questionList.length = {0}", questionList.Count);

            List<List<string>> newQuestionString = new List<List<string>>();
            List<List<string>> newAnswerString = new List<List<string>>();

            for (int i = 0; i < num; i++)
            {
                List<string> strings = new List<string>();
                //Console.WriteLine("numIndex[{0}-1]= {1}",i, numIndex[i] - 1   );
                //Console.WriteLine("questionList[numIndex[{0}]-1].Count = {1}", i, questionList[numIndex[i] - 1].Count);
                for (int j = 0; j < questionList[numIndex[i] - 1].Count; j++)
                {
                    strings.Add(questionList[numIndex[i] - 1][j]);
                }
                newQuestionString.Add(strings);
            }

            for (int i = 0; i < num; i++)
            {
                List<string> strings = new List<string>();
                //Console.WriteLine("numIndex[{0}-1]= {1}",i, numIndex[i] - 1   );
                //Console.WriteLine("questionList[numIndex[{0}]-1].Count = {1}", i, questionList[numIndex[i] - 1].Count);
                for (int j = 0; j < answerList[numIndex[i] - 1].Count; j++)
                {
                    strings.Add(answerList[numIndex[i] - 1][j]);
                }
                newAnswerString.Add(strings);
            }

            return (newQuestionString, newAnswerString);
        }

        /// <summary>
        /// 根据给定的题库路径搜索题库文件，并按照给定的抓取数量，随机抓取补不重复的题库，并形成一个数组
        /// </summary>
        /// <param name="questionStoragePath">题库路径</param>
        /// <param name="num">该题库选题题数</param>
        /// <returns>选中的题目顺序数组</returns>
        internal static int[] GetNeededQuestionAndAnswerIndexFromQuestionStorage(string questionStoragePath, int num)
        {
            int[] numIndex = new int[num];
            //Console.WriteLine("questionStoragePath:{0}", questionStoragePath);
            if (!File.Exists(questionStoragePath)) return numIndex; ///检查路径文件存在
            StreamReader streamReader = File.OpenText(questionStoragePath);//打开文件

            string nextLine;
            int lineIndex = 0;//行号计数器

            List<List<string>> questionList = new List<List<string>>();
            List<List<string>> answerList = new List<List<string>>();

            List<int> questionIndexList = new List<int>();
            List<int> answerIndexList = new List<int>();

            nextLine = streamReader.ReadLine();//nextline 读取第一行

            lineIndex++;//总行号计数器+1

            while (nextLine != null)
            {
                //string templine = nextLine;
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目："))
                {
                    questionIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        list.Add(nextLine);
                    }
                    questionList.Add(list);
                }
                else if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案："))
                {
                    answerIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        //Console.WriteLine("nextLine:{0}", nextLine);
                        //Console.WriteLine("nextLine.Substring(0, 3):{0}", nextLine.Substring(0, 3));
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                        if (nextLine != "") list.Add(nextLine);//答案不要空行
                    }
                    answerList.Add(list);
                }
                else//对于没有出现答案和题目的情况，如下处理
                {
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if ((nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        if ((nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                    }
                }
            }
            streamReader.Close();

            if (questionIndexList.Count == 0 || answerIndexList.Count == 0) return numIndex;//表明里面没有数据内容，故不符合题库条件

            Random random = new Random();
            int num_random;
            int counter = 0;
            while (counter < num)
            {
                num_random = random.Next(1, questionList.Count + 1);
                if (!numIndex.Contains<int>(num_random))
                {
                    numIndex[counter] = num_random;
                    counter++;
                }
            }
            Array.Sort(numIndex);
            return numIndex;
        }

        /// <summary>
        /// 根据给定的题库路径和索引数组，确定题库内容并返回
        /// </summary>
        /// <param name="questionStoragePath">题库路径</param>
        /// <param name="index">索引</param>
        /// <returns>题目内容</returns>
        internal static List<List<string>> GetNeededQuestionFromQuestionStorageByIndex(string questionStoragePath, int[] index)
        {
            //Console.WriteLine("questionStoragePath:{0}", questionStoragePath);
            if (!File.Exists(questionStoragePath)) return null;///检查路径文件存在
            StreamReader streamReader = File.OpenText(questionStoragePath);//打开文件

            string nextLine;
            int lineIndex = 0;//行号计数器

            List<List<string>> questionList = new List<List<string>>();

            List<int> questionIndexList = new List<int>();

            nextLine = streamReader.ReadLine();//nextline 读取第一行

            lineIndex++;//总行号计数器+1

            while (nextLine != null)
            {
                //string templine = nextLine;
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目："))
                {
                    questionIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        list.Add(nextLine);
                    }
                    questionList.Add(list);
                }
                else if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案："))
                {
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                    }
                }
                else//对于没有出现答案和题目的情况，如下处理
                {
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if ((nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        if ((nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                    }
                }
            }
            streamReader.Close();

            if (questionIndexList.Count == 0) return null;//表明里面没有数据内容，故不符合题库条件

            List<List<string>> newQuestionString = new List<List<string>>();

            for (int i = 0; i < index.Length; i++)
            {
                List<string> strings = new List<string>();
                for (int j = 0; j < questionList[index[i] - 1].Count; j++)
                {
                    strings.Add(questionList[index[i] - 1][j]);
                }
                newQuestionString.Add(strings);
            }
            return newQuestionString;
        }

        /// <summary>
        /// 根据给定的题库路径和索引数组，确定题库答案内容并返回
        /// </summary>
        /// <param name="questionStoragePath">题库路径</param>
        /// <param name="index">索引</param>
        /// <returns>答案内容</returns>
        internal static List<List<string>> GetNeededAnswerFromQuestionStorageByIndex(string questionStoragePath, int[] index)
        {
            //Console.WriteLine("questionStoragePath:{0}", questionStoragePath);
            if (!File.Exists(questionStoragePath)) return null;///检查路径文件存在
            StreamReader streamReader = File.OpenText(questionStoragePath);//打开文件

            string nextLine;
            int lineIndex = 0;//行号计数器

            List<List<string>> answerList = new List<List<string>>();

            List<int> answerIndexList = new List<int>();

            nextLine = streamReader.ReadLine();//nextline 读取第一行

            lineIndex++;//总行号计数器+1

            while (nextLine != null)
            {
                //string templine = nextLine;
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目："))
                {
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                    }
                }
                else if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案："))
                {
                    answerIndexList.Add(lineIndex);
                    List<string> list = new List<string>();
                    list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        //Console.WriteLine("nextLine:{0}", nextLine);
                        //Console.WriteLine("nextLine.Substring(0, 3):{0}", nextLine.Substring(0, 3));
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                        if (nextLine != "") list.Add(nextLine);//答案不要空行
                    }
                    answerList.Add(list);
                }
                else//对于没有出现答案和题目的情况，如下处理
                {
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        if ((nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        if ((nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                    }
                }
            }
            streamReader.Close();

            if (answerIndexList.Count == 0) return null;//表明里面没有数据内容，故不符合题库条件

            //Console.WriteLine("numIndex.length = {0}", index.Length);
            //Console.WriteLine("questionList.length = {0}", questionList.Count);

            List<List<string>> newAnswerString = new List<List<string>>();

            for (int i = 0; i < index.Length; i++)
            {
                List<string> strings = new List<string>();
                //Console.WriteLine("numIndex[{0}-1]= {1}",i, numIndex[i] - 1   );
                //Console.WriteLine("questionList[numIndex[{0}]-1].Count = {1}", i, questionList[numIndex[i] - 1].Count);
                for (int j = 0; j < answerList[index[i] - 1].Count; j++)
                {
                    strings.Add(answerList[index[i] - 1][j]);
                }
                newAnswerString.Add(strings);
            }
            return newAnswerString;
        }

        /// <summary>
        /// 连接两个List数据内容
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        internal static List<List<string>> CombineList(List<List<string>> a, List<List<string>> b)
        {
            List<List<string>> newOne = new List<List<string>>();
            if (a != null)
            {
                for (int i = 0; i < a.Count; i++)
                {
                    List<string> strings = new List<string>();
                    for (int j = 0; j < a[i].Count; j++)
                    {
                        strings.Add(a[i][j]);
                    }
                    newOne.Add(strings);
                }
            }

            if (b != null)
            {
                for (int i = 0; i < b.Count; i++)
                {
                    List<string> strings = new List<string>();
                    for (int j = 0; j < b[i].Count; j++)
                    {
                        strings.Add(b[i][j]);
                    }
                    newOne.Add(strings);
                }
            }

            return newOne;
        }

        /// <summary>
        /// 阿拉伯数字转换为大写数字并加上顿号
        /// </summary>
        /// <param name="num">阿拉伯数字</param>
        /// <returns>返回大写数字加顿号</returns>
        public static string IntToBigNum(int num)
        {
            switch (num)
            {
                case 1: return "一、";
                case 2: return "二、";
                case 3: return "三、";
                case 4: return "四、";
                case 5: return "五、";
                case 6: return "六、";
                case 7: return "七、";
                case 8: return "八、";
                case 9: return "九、";
                case 10: return "十、";
                case 11: return "十一、";
                case 12: return "十二、";
                case 13: return "十三、";
                case 14: return "十四、";
                case 15: return "十五、";
                case 16: return "十六、";
                case 17: return "十七、";
                case 18: return "十八、";
                case 19: return "十九、";
                default: return "〇、";
            }
        }

        public static float GetScreenScaling(NewExamTaskContentForm form)
        {
            float dpiX;
            Graphics graphics = form.CreateGraphics();
            dpiX = graphics.DpiX;
            //Console.WriteLine(dpiX.ToString());

            if (dpiX == 96) { return 1; }
            else if (dpiX == 120) { return 1.25f; }
            else if (dpiX == 144) { return 1.5f; }
            else if (dpiX == 192) { return 2f; }
            else { return dpiX / 96; }
        }

        public static float GetScreenScaling(EditExamTaskContentForm form)
        {
            float dpiX;
            Graphics graphics = form.CreateGraphics();
            dpiX = graphics.DpiX;
            //Console.WriteLine(dpiX.ToString());

            if (dpiX == 96) { return 1; }
            else if (dpiX == 120) { return 1.25f; }
            else if (dpiX == 144) { return 1.5f; }
            else if (dpiX == 192) { return 2f; }
            else { return dpiX / 96; }
        }
    }
}