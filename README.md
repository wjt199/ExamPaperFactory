# ExamPaperFactory

一、软件介绍

  通过试卷任务文件描述试卷生成方法（将任务文件加载至本软件后，可以生成相对应的试卷格式和内容）
  
二、使用本软件限制条件（条件比较苛刻，能力有限，望理解）

（一）Windows7及以上版本的64位系统；

（二）系统安装有“.Net Framework 4.7.2”及以上环境；

（三）系统安装有Office2016及以上版本办公软件，且“docx”文件默认打开方式设置为Office（若为WPS会出现难以预料的问题）。

三、使用方法

（一）试卷任务文件格式为“xml”，建议使用本软件生成，不建议使用文本文档软件修改“xml”任务文件，防止文件因格式问题导致错误；

（二）试卷的单位姓名部分已写死，格式为： “单位：___________    姓名：___________    成绩：___________”， 如果不满意可以在试卷文件生成后进行手动更改；

（三）题库路径选择方法为：在题库框中使用鼠标双击，在弹出的对话框中选择对应的题库路径；

（四）题库文件设置要求：

   1.题库文件格式为“txt”文本文件，题库名称无限制；
    
   2.题库文本组成以“题目+答案”的循环方式组成（每个题目体感部分结束后换行，紧跟着对应题目的答案内容部分），题库文件至少含1个题目和1个答案；
   
   3.“题目”格式：以“题目：”或“题目:”开头，后面为题目本体，可以多行描述，以“答案：”开头的下一行作为“题目”内容结束的标志；
   
   4.“答案”格式：以“答案：”或“答案:”开头，后面为答案本体，可以多行描述，以“题目：”开头的下一行或没有下一行作为“答案”内容结束的标志；
   
   5.题库文件内容不带序号。
   
四、其他

   在使用过程中存在疑问、发现软件的BUG或有更好的意见建议，欢迎联系作者（E-mail：wjt199@sina.com）,感激不尽！
