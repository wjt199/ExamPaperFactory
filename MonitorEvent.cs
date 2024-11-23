using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MonitorEvent
{
    /// <summary>
    /// 定义委托
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    public delegate void MonitorEventHandler(object sender, EventArgs e);

    /// <summary>
    /// 时间数据类
    /// </summary>
    public class MsgEventArgs : EventArgs
    {
        /// <summary>
        /// 修改时间
        /// </summary>
        public DateTime ChanggeTime { get; set; }

        /// <summary>
        /// 提示信息
        /// </summary>
        public string ToSend { get; set; }
        public MsgEventArgs(DateTime changgeTime, string toSend)
        {
            this.ChanggeTime = changgeTime;
            this.ToSend = toSend;
        }
    }

    /// <summary>
    /// 事件订阅者
    /// </summary>
    public class MonitorText
    {
        public string name = "文本文档";
        //定义监控文本事件
        public event MonitorEventHandler MonitorEvent;
        //上次文件更新时间用于判断文件是否修改过
        private DateTime _lastWriteTime = File.GetLastWriteTime(@"C:\Users\Administrator\Desktop\1.txt");
        public MonitorText()
        {

        }
        // 文件更新调用
        protected virtual void OnTextChange(MsgEventArgs e)
        {
            if (MonitorEvent != null)
            {
                //不为空，处理事件
                MonitorEvent(this, e);
            }
        }

        //事件监听的方法
        public void BeginMonitor()
        {
            DateTime bCurrentTime;

            while (true)
            {
                bCurrentTime = File.GetLastWriteTime(@"C:\Users\Administrator\Desktop\1.txt");
                if (bCurrentTime != _lastWriteTime)
                {
                    _lastWriteTime = bCurrentTime;
                    MsgEventArgs msg = new MsgEventArgs(bCurrentTime, "文本改变了");
                    OnTextChange(msg);
                }
                //0.1秒监控一次
                Thread.Sleep(100);
            }
        }

    }

    /// <summary>
    /// 事件监听者
    /// </summary>
    public class Administrator
    {
        //管理员事件处理方法
        public void OnTextChange(object Sender, EventArgs e)
        {
            MonitorText monitorText = (MonitorText)Sender;
            Console.WriteLine("尊敬的管理员：" + DateTime.Now.ToString() + ": " + monitorText.name + "发生改变.");
        }
    }

    class Program
    {
        static MonitorText MonitorTextEventSource;
        static void Main(string[] args)
        {
            MonitorTextEventSource = new MonitorText();
            //1. 启动后台线程添加监视事件
            var thrd = new Thread(MonitorTextEventSource.BeginMonitor);
            thrd.IsBackground = true;
            thrd.Start();
            //2实例化管理员类
            Administrator ad = new Administrator();
            //4订阅事件
            MonitorTextEventSource.MonitorEvent += ad.OnTextChange;
            Console.ReadLine();
        }
    }
}