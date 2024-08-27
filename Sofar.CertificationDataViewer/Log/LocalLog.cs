using System;
using System.IO;
using System.Text;

namespace ATE.Business
{
    /// <summary>
    ///  本地日志
    /// </summary>
    internal class LocalLog
    {
        /// <summary>
        /// 单实例
        /// </summary>
        private static LocalLog _instance = null;

        /// <summary>
        /// 静态锁
        /// </summary>
        private static readonly object obj = new object();

        /// <summary>
        /// 写入流
        /// </summary>
        private StreamWriter _streamWrite = null;

        /// <summary>
        /// 本地日志存放目录
        /// </summary>
        public string LogDirectory = string.Empty;

        public static LocalLog Instance
        {
            get
            {
                if (null == _instance)
                {
                    lock (obj)
                    {
                        if (_instance == null)
                        {
                            _instance = new LocalLog();
                        }
                    }
                }
                return _instance;
            }
        }

        private LocalLog()
        {
            try
            {
                string datetime = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                LogDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\logs\\log-" + datetime;

                Directory.CreateDirectory(LogDirectory);

                string fileName = string.Format("log-{0}.txt", datetime);
                string path = LogDirectory + "\\" + fileName;

                // 程序日志输出
                _streamWrite = new StreamWriter(path, true, Encoding.UTF8);
                _streamWrite.AutoFlush = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public void Write(string content)
        {
            try
            {
                _streamWrite?.WriteLine(content);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}