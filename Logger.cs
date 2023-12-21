using System;
using System.IO;

namespace SimpleOfficeCreator.Standard
{
    public class Logger
    {
        private Logger() { }
        //private static 인스턴스 객체
        private static readonly Lazy<Logger> _instance = new Lazy<Logger>(() => new Logger());
        //public static 의 객체반환 함수
        public static Logger Instance { get { return _instance.Value; } }

        private bool lastLogIsContainLineBreak = true;
        public string Write(string log = "", bool linebreak = true, int minBlank = 0, ConsoleColor color = ConsoleColor.White)
        {
            string path = Path.Combine(Environment.CurrentDirectory, "Log");
            if (Directory.Exists(path))
            {
                string configFile = Path.Combine(path, "officelog.enable");
                if (File.Exists(configFile))
                {
                    string strLogFileName = string.Format("officelog_{0}.txt", DateTime.Now.ToString("yyyyMMdd_HH"));
                    string strLogPath = Path.Combine(path, strLogFileName);

                    using (StreamWriter fileWrite = new StreamWriter(strLogPath, true))
                    {
                        if (lastLogIsContainLineBreak)
                        {
                            string timeLog = DateTime.Now.ToString("HH:mm:ss fff");
                            fileWrite.Write($"[{timeLog}] ");
                        }

                        if (linebreak)
                        {
                            if (minBlank == 0)
                            {
                                fileWrite.WriteLine(log);
                            }
                            else
                            {
                                fileWrite.WriteLine(String.Format("{0,-" + minBlank + "}", log));
                            }
                        }
                        else
                        {
                            if (minBlank == 0)
                            {
                                fileWrite.Write(log);
                            }
                            else
                            {
                                fileWrite.Write(String.Format("{0,-" + minBlank + "}", log));
                            }
                        }

                        lastLogIsContainLineBreak = linebreak;
                    }
                }
            }
            Console.ForegroundColor = color;
            if (linebreak)
            {
                if (minBlank == 0)
                {
                    Console.WriteLine(log);
                }
                else
                {
                    Console.WriteLine(String.Format("{0,-" + minBlank + "}", log));
                }
            }
            else
            {
                if (minBlank == 0)
                {
                    Console.Write(log);
                }
                else
                {
                    Console.Write(String.Format("{0,-" + minBlank + "}", log));
                }
            }
            Console.ForegroundColor = ConsoleColor.White;

            return log;
        }

        public void WriteWarning(string log)
        {
            Write(log, true, 0, ConsoleColor.Yellow);
        }
        public void WriteError(string log)
        {
            Write(log, true, 0, ConsoleColor.Red);
        }
    }
}
