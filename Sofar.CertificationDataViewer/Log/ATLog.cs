using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;

namespace ATE.Business
{
    public enum EnumLog
    {
        [Description("Data")]
        D,

        [Description("Error")]
        E,
    }

    public enum EnumLogTag
    {
        [Description("普通日志")]
        Log,

        [Description("老化日志")]
        Aging,
    }

    public class ATLog
    {
        /// <summary>
        /// 日志级别
        /// </summary>
        public static EnumLog[] Levels = new EnumLog[] { EnumLog.E, EnumLog.D };

        static ATLog()
        {
            // 删除本地日志文件
            try
            {
                DateTime nowTime = DateTime.Now;
                DirectoryInfo root = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "\\logs");
                DirectoryInfo[] dics = root.GetDirectories();//获取文件夹

                FileAttributes attr = File.GetAttributes(AppDomain.CurrentDomain.BaseDirectory + "\\logs");
                if (attr == FileAttributes.Directory)//判断是不是文件夹
                {
                    foreach (DirectoryInfo file in dics)//遍历文件夹
                    {
                        TimeSpan t = nowTime - file.CreationTime;  //当前时间  减去 文件创建时间
                        int day = t.Days;
                        if (day >= 30)   //保存的时间 ；  单位：天
                        {
                            Directory.Delete(file.FullName, true);  //删除超过时间的文件夹
                        }
                    }
                }
            }
            catch
            {
            }
            finally
            {
                Thread.Sleep(20);
            }
        }

        /// <summary>
        /// 记录Data 数据
        /// </summary>
        /// <param name="msg"></param>
        public static void DLog(string msg)
        {
            if (Levels.Contains(EnumLog.D))
                ConsoleLog(EnumLog.D, EnumLogTag.Log, msg);
        }

        /// <summary>
        /// 记录Data 数据
        /// </summary>
        /// <param name="msg"></param>
        public static void ELog(string msg)
        {
            if (Levels.Contains(EnumLog.E))
                ConsoleLog(EnumLog.E, EnumLogTag.Log, msg);
        }

        /// <summary>
        /// 输出日志
        /// </summary>
        /// <param name="level">日志级别</param>
        /// <param name="tag">日志标记</param>
        /// <param name="msg">日志信息</param>
        private static void ConsoleLog(EnumLog level, EnumLogTag tag, string msg)
        {
            string logInfo = $"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff")} {level}/{tag} {msg}";
            Console.WriteLine(logInfo);
            LocalLog.Instance.Write(logInfo);
        }
    }
}