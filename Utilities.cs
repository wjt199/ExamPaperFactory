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
        /// <summary> *不再使用*
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
        /// 功能：将数字转换为从1到该数字组成的字符串数组
        /// </summary>
        /// <param name="num">大于1的整数</param>
        /// <returns>字符串数组</returns>
        public static String[] NumToStringList(int num)
        {
            List<string> strList = new List<String>();
            for (int i = 1; i <= num; i++) strList.Add(i.ToString());
            return strList.ToArray();
        }

        /// <summary>
        /// 功能：将两个数字之间的所有整数转换为由小到大的字符串数组
        /// </summary>
        /// <param name="startNum">起始数字（int）</param>
        /// <param name="endNum">结束数字（int）</param>
        /// <returns>字符串数组</returns>
        public static String[] NumToStringList(int startNum, int endNum)
        {
            List<string> strList = new List<String>();
            for (int i = startNum; i <= endNum; i++) strList.Add(i.ToString());
            return strList.ToArray();
        }

        /// <summary>
        /// 功能：将两个数字之间的数字按照给定的步进值转换为由小到大的字符串数组
        /// </summary>
        /// <param name="startNum">起始数字（float）</param>
        /// <param name="endNum">结束数字（float）</param>
        /// <param name="step">步进值（float）</param>
        /// <returns>字符串数组</returns>
        public static String[] NumToStringList(float startNum, float endNum, float step)
        {
            List<string> strList = new List<String>();
            for (float i = startNum; i <= endNum; i += step) strList.Add(i.ToString());
            return strList.ToArray();
        }

        /// <summary>
        /// 功能：根据传入的单一题库绝对路径读取题库文件，并统计题库中所有题目的数量，如果题目数和答案数不同则返回0
        /// 实现方法：分两次打开路径文件，逐行读取，分别按照题目和答案进行计数，如果两个数量一致则返回该数，否则返回0
        /// </summary>
        /// <param name="questionStoragePath">单一题库绝对路径</param>
        /// <returns>正常情况返回：题库对应题目总数；题库文件不存在或题目与答案数量不同返回：0</returns>
        public static int CountQuestionNumFromQuestionStoragePath(string questionStoragePath)
        {
            //题目计数器初始值为0
            int numQuestion = 0;
            //答案计数器初始值为0
            int numAnswer = 0;

            //当前读取的字符串
            string nextLine;

            //如果对应的题库路径给定的文件不存在，返回0
            if (!File.Exists(questionStoragePath)) return numQuestion;

            StreamReader streamReaderQuestion = File.OpenText(questionStoragePath);
            //循环读取题库文件，找到所有行首带有“题目:”的行，对其进行统计
            while ((nextLine = streamReaderQuestion.ReadLine()) != null)
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) numQuestion++;
            streamReaderQuestion.Close();

            StreamReader streamReaderAnswer = File.OpenText(questionStoragePath);
            //循环读取题库文件，找到所有行首带有“答案:”的行，对其进行统计
            while ((nextLine = streamReaderAnswer.ReadLine()) != null)
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) numAnswer++;
            streamReaderAnswer.Close();

            return numQuestion == numAnswer ? numAnswer : 0;
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
            //检查路径文件存在
            if (!File.Exists(questionStoragePath)) return false;
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
                    if (nextLine.Substring(3) != String.Empty) list.Add(nextLine.Substring(3));//在除去“题目：”后，后面还有内容就加进去，否则就不用加了
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
                    if (nextLine.Substring(3) != String.Empty) list.Add(nextLine.Substring(3));//在除去“答案：”后，后面还有内容就加进去，否则就不用加了
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        lineIndex++;//总行号计数器+1
                        //Console.WriteLine("nextLine:{0}", nextLine);
                        //Console.WriteLine("nextLine.Substring(0, 3):{0}", nextLine.Substring(0, 3));
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                        if (nextLine != String.Empty) list.Add(nextLine);//答案不要空行
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
        /// 功能：根据给定的题库路径读取对应的题库，按照给定的选题数量，在题库中随机抽取题库，并将抽中的题目和对应的答案列表以列表元组的方式返回
        /// 实现方法：根据题库路径调用本工具类下的CountQuestionNumFromQuestionStoragePath()方法计算题目总数，通过随机数在对应的范围内循环生成不重复的随机数，
        /// 并将所有的随机数加入数组，全部加入完成后，对数组进行排序，并数组时进行内容检查，如果包含0，返回(null,null)，全过程严格检查。
        /// 根据随机数组将题目和对应的答案分别提取形成类型为List<List<string>>的列表，将此两个列表以元组(List<List<string>>, List<List<string>>)的形式返回。
        /// </summary>
        /// <param name="questionStoragePath">单一题库绝对路径</param>
        /// <param name="amount">在指定题库中选题题数</param>
        /// <returns>正常情况：Item1:选中的题库内容;Item2:对应的答案内容；路径不存在或路径对应题库内容不包含题目和答案时均返回：(null,null)</returns>
        internal static (List<List<string>>, List<List<string>>) GetNeededQuestionAndAnswerFromQuestionStorage(string questionStoragePath, int amount)
        {
            //通过函数获取题库路径对应的题目总数
            int questionOrAnswerNum = CountQuestionNumFromQuestionStoragePath(questionStoragePath);
            //路径文件不存在、题目数量为0或选题数小于选题数时返回(null, null)
            if (questionOrAnswerNum == 0 || questionOrAnswerNum < amount) return (null, null);

            //打开读取题库路径
            StreamReader streamReader = File.OpenText(questionStoragePath);

            //"是否存在“题目：”内容"初始值为不存在
            bool isQuestionExist = false;
            //"是否存在“答案：”内容"初始值为不存在
            bool isAnswerExist = false;

            //全部题目列表
            List<List<string>> allQuestionList = new List<List<string>>();
            //全部答案列表
            List<List<string>> allAnswerList = new List<List<string>>();
            //选中的题库内容列表
            List<List<string>> selectedQuestionList = new List<List<string>>();
            //选中的答案列表
            List<List<string>> selectedAnswerList = new List<List<string>>();

            //nextline 读取第一行
            string nextLine = streamReader.ReadLine();
            //按照规则循环处理下一行字符串内容，直到处理结束
            while (nextLine != null)
            {
                //找到带有“题目：”行
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目："))
                {
                    if(!isQuestionExist) isQuestionExist = true;
                    List<string> list = new List<string>();
                    //在除去“题目：”后，还有内容就加入list
                    if (nextLine.Substring(3) != String.Empty) list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        //题目内容以下一行出现“答案："作为结束标志
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break;
                        //题目内容剔除空行
                        if (nextLine != String.Empty) list.Add(nextLine);
                    }
                    allQuestionList.Add(list);
                }
                else if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案："))
                {
                    if (!isAnswerExist) isAnswerExist = true;
                    List<string> list = new List<string>();
                    //在除去“答案：”后，还有内容就加入list
                    if (nextLine.Substring(3) != String.Empty) list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        //答案内容以下一行出现“题目："作为结束标志
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break;
                        //答案内容剔除空行
                        if (nextLine != String.Empty) list.Add(nextLine);
                    }
                    allAnswerList.Add(list);
                }
                else
                {
                    //若该行既没有“题目：”也没有“答案：”，则循环读取下一行，直到行首出现“题目：”或“答案：”时停止跳出循环
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break; //遇到答案就退出本循环
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break; //遇到题目就退出本循环
                    }
                }
            }
            streamReader.Close();
                        
            //题库文件里没有内容，返回(null, null)
            if (!isQuestionExist || !isAnswerExist) return (null, null);

            int[] numIndex = new int[amount];
            Random random = new Random();
            int numRandom;
            int counter = 0;
            while (counter < amount)
            {
                //随机数范围：[1,题目总数]
                numRandom = random.Next(1, questionOrAnswerNum + 1);
                //检查数组中有没有上面抽到的数字，如果有相同的数字
                if (!numIndex.Contains<int>(numRandom))
                {
                    numIndex[counter] = numRandom;
                    counter++;
                }
            }
            //对选中的题目数组进行排序
            Array.Sort(numIndex);
            //如果数组中包含0，返回(null, null)（通常情况下不会出问题）
            if (numIndex.Contains<int>(0)) return (null, null);

            for (int i = 0; i < amount; i++)
            {
                List<string> strings = new List<string>();
                foreach (string str in allQuestionList[numIndex[i] - 1]) strings.Add(str);
                selectedQuestionList.Add(strings);
            }

            for (int i = 0; i < amount; i++)
            {
                List<string> strings = new List<string>();
                foreach (string str in allAnswerList[numIndex[i] - 1]) strings.Add(str);
                selectedAnswerList.Add(strings);
            }

            return (selectedQuestionList, selectedAnswerList);
        }

        /// <summary>
        /// 功能：根据给定的题库路径读取对应的题库，按照给定的选题数量，在题库中随机抽取题目，并将抽取的题库排序按从小到大的顺序排列后形成数组作为结果返回
        /// 实现方法：根据题库路径调用本工具类下的CountQuestionNumFromQuestionStoragePath()方法计算题目总数，通过随机数在对应的范围内循环生成不重复的随机数，
        /// 并将所有的随机数加入数组，全部加入完成后，对数组进行排序，最后返回数组时进行内容检查，如果包含0，依然返回null，全过程严格检查。
        /// </summary>
        /// <param name="questionStoragePath">单一题库绝对路径</param>
        /// <param name="amount">在指定题库中选题题数</param>
        /// <returns>正常情况返回：被选中的题目排序数组（题目序号从1开始）；题库文件不存在、题库文件里面没有题目或题目的总数小于选题数时返回：空数组</returns>
        internal static int[] GetNeededQuestionAndAnswerIndexFromQuestionStorage(string questionStoragePath, int amount)
        {
            //获取题库中题目的数量，如果题目数量为0、题库有误或数量有误均返回0
            int questionNum = CountQuestionNumFromQuestionStoragePath(questionStoragePath);
            //题目数量为0或选题数小于选题数时返回null
            if (questionNum == 0 || questionNum<amount) return null;

            //新建需要返回的数组，初始值为空
            int[] numIndex = new int[amount];
            //创建随机数对象
            Random random = new Random();
            //伪随机数
            int numRandom;
            //计数器
            int counter = 0;
            while (counter < amount)
            {
                //随机数范围：[1,题目总数]
                numRandom = random.Next(1, questionNum + 1);
                //检查数组中有没有上面抽到的数字，如果有相同的数字
                if (!numIndex.Contains<int>(numRandom))
                {
                    numIndex[counter] = numRandom;
                    counter++;
                }
            }
            //按从小到大的顺序对随机数进行排序
            Array.Sort(numIndex);
            //如果数组中不包含0，就返回该数组，否则返回null（通常情况下不会出问题）
            return numIndex.Contains<int>(0) ? null : numIndex;
        }

        /// <summary>
        /// 功能：根据单一题库绝对路径确定题库内容，按照索引数组在题库内容里面寻找对应的题目位置，将选择的题库题目内容整合并以类型为List<List<string>>的形式返回。
        /// 实现方法：先将题库里全部题目内容提取形成一个类型为List<List<string>>的列表，后根据索引数组里的序号挨个提取对应题目内容，整合为一个新的类型为List<List<string>>的列表，作为输出返回
        /// </summary>
        /// <param name="questionStoragePath">单一题库绝对路径</param>
        /// <param name="index">索引数组：索引内容起始值为1非0（注意！）</param>
        /// <returns>正常情况返回：题目内容列表；路径不存在或路径对应题库内容不包含题目时均返回：null</returns>
        internal static List<List<string>> GetNeededQuestionFromQuestionStorageByIndex(string questionStoragePath, int[] index)
        {
            //检查路径文件存在情况，不存在直接返回null
            if (!File.Exists(questionStoragePath)) return null;
            //打开读取题库路径
            StreamReader streamReader = File.OpenText(questionStoragePath);
            //全部题目列表
            List<List<string>> questionList = new List<List<string>>();
            //索引题目列表
            List<List<string>> selectedQuestionList = new List<List<string>>();
            //答案内容存在标签
            bool isQuestionExist = false;
            //nextline 读取第一行
            string nextLine = streamReader.ReadLine();
            //循环读取题库路径内容
            while (nextLine != null)
            {
                //找到带有“题目：”行
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目："))
                {
                    if (!isQuestionExist) isQuestionExist = true;
                    List<string> list = new List<string>();
                    //在除去“题目：”后，还有内容就加入list
                    if (nextLine.Substring(3) != String.Empty) list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        //题目内容以下一行出现“答案："作为结束标志
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break;
                        //题目内容剔除空行
                        if (nextLine != String.Empty) list.Add(nextLine);
                    }
                    questionList.Add(list);
                }
                else
                {
                    //若该行没有“题目：”，则循环读取下一行，直到行首出现“题目：”时停止跳出循环
                    while ((nextLine = streamReader.ReadLine()) != null)
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break;
                }
            }
            streamReader.Close();

            //循环完后没有“题目:”内容，返回null
            if (!isQuestionExist) return null;
            //根据数组内数字要求循环对selectedQuestionList列表填充选定题库内容
            foreach (int num in index)
            {
                List<string> strings = new List<string>();
                foreach (string str in questionList[num - 1]) strings.Add(str);
                selectedQuestionList.Add(strings);
            }
            //清空临时数据，节约内存
            questionList.Clear();
            return selectedQuestionList;
        }

        /// <summary>
        /// 功能：根据单一题库绝对路径确定题库内容，按照索引数组在题库内容里面寻找对应的答案位置，将选择的题库答案内容整合并以类型为List<List<string>>的形式返回。
        /// 实现方法：先将题库里全部答案内容提取形成一个类型为List<List<string>>的列表，后根据索引数组里的序号挨个提取对应答案内容，整合为一个新的类型为List<List<string>>的列表，作为输出返回
        /// </summary>
        /// <param name="questionStoragePath">单一题库绝对路径</param>
        /// <param name="index">索引数组：索引内容起始值为1非0（注意！）</param>
        /// <returns>正常情况返回：答案内容列表；路径不存在或路径对应题库内容不包含答案时均返回：null</returns>
        internal static List<List<string>> GetNeededAnswerFromQuestionStorageByIndex(string questionStoragePath, int[] index)
        {
            //检查路径文件存在情况，不存在直接返回null
            if (!File.Exists(questionStoragePath)) return null;
            //打开读取题库路径
            StreamReader streamReader = File.OpenText(questionStoragePath);
            //全部答案列表
            List<List<string>> answerList = new List<List<string>>();
            //索引答案列表
            List<List<string>> selectedAnswerList = new List<List<string>>();
            //答案内容存在标签
            bool isAnswerExist = false;
            //nextline 读取第一行
            string nextLine = streamReader.ReadLine();
            //循环读取题库路径内容
            while (nextLine != null)
            {
                //找到带有“答案：”行
                if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案："))
                {
                    if (!isAnswerExist) isAnswerExist = true;
                    List<string> list = new List<string>();
                    //在除去“答案：”后，还有内容就加入list
                    if (nextLine.Substring(3) != String.Empty) list.Add(nextLine.Substring(3));
                    while ((nextLine = streamReader.ReadLine()) != null)
                    {
                        //答案内容以下一行出现“题目："作为结束标志
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "题目:" || nextLine.Substring(0, 3) == "题目：")) break;
                        //答案内容剔除空行
                        if (nextLine != String.Empty) list.Add(nextLine);
                    }
                    answerList.Add(list);
                }
                else
                {
                    //若该行没有“答案：”，则循环读取下一行，直到行首出现“答案：”时停止
                    while ((nextLine = streamReader.ReadLine()) != null)
                        if (nextLine.Length >= 3 && (nextLine.Substring(0, 3) == "答案:" || nextLine.Substring(0, 3) == "答案：")) break;
                }
            }
            streamReader.Close();

            //循环完后没有“答案:”内容，返回null
            if (!isAnswerExist) return null;
            //根据数组内数字要求循环对selectedAnswerList列表填充选定答案内容
            foreach (int num in index)
            {
                List<string> strings = new List<string>();
                foreach (string str in answerList[num - 1]) strings.Add(str);
                selectedAnswerList.Add(strings);
            }
            //清空临时数据，节约内存
            answerList.Clear();
            return selectedAnswerList;
        }

        /// <summary>
        /// 功能：连接两个类型为List<List<string>>的数据内容。
        /// 实现方法：新建一个类型为List<List<string>>的变量，将aList中的数据分解后按顺序添加进去,再将bList中的数据分解后添加进去，如果为空就跳过不添加
        /// </summary>
        /// <param name="aList">List数据A</param>
        /// <param name="bList">List数据B</param>
        /// <returns></returns>
        internal static List<List<string>> CombineList(List<List<string>> aList, List<List<string>> bList)
        {
            List<List<string>> newList = new List<List<string>>();

            //添加aList数据
            if (aList != null)
            {
                foreach (List<string> list in aList)
                {
                    List<string> strings = new List<string>();
                    foreach (string str in list) strings.Add(str);
                    newList.Add(strings);
                }
            }
            //添加bList数据
            if (bList != null)
            {
                foreach (List<string> list in bList)
                {
                    List<string> strings = new List<string>();
                    foreach (string str in list) strings.Add(str);
                    newList.Add(strings);
                }
            }

            return newList;
        }

        /// <summary>
        /// 功能：将阿拉伯数字转换为对应的大写数字并加上顿号，范围：1-19，超出范围一律按0处理
        /// </summary>
        /// <param name="num">阿拉伯数字(形如：5)</param>
        /// <returns>数字大写加顿号（形如：五、）</returns>
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

        /// <summary>
        /// 功能：获取当前屏幕的缩放系数
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static float GetScreenScaling(NewExamTaskContentForm form)
        {
            Graphics graphics = form.CreateGraphics();
            switch (graphics.DpiX)
            {
                case 96: return 1;
                case 120: return 1.25f;
                case 144: return 1.5f;
                case 192: return 2f;
                default: return graphics.DpiX / 96;
            }
        }

        /// <summary>
        /// 功能：获取当前屏幕的缩放系数
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static float GetScreenScaling(EditExamTaskContentForm form)
        {
            Graphics graphics = form.CreateGraphics();
            switch (graphics.DpiX)
            {
                case 96: return 1;
                case 120: return 1.25f;
                case 144: return 1.5f;
                case 192: return 2f;
                default: return graphics.DpiX / 96;
            }
        }
    }
}
