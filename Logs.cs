using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace ConsoleApplication1
{
    class Logs
    {
        public static void Write(string line)
        {
            try
            {
                string curDir = Application.StartupPath + @"\logs\";
                string logFile = @"\Test.log";
                StreamWriter lg = new StreamWriter(curDir + logFile, true);
                Console.WriteLine(DateTime.Now + " [INF] " + line);
                lg.WriteLine(DateTime.Now + " [INF] " + line);
                lg.Close();
                lg.Dispose();
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }   
        }

        public static void CreateLogFile()
        {
            string curDir = Application.StartupPath + @"\logs\";
            string logFile = @"\Test.log";

            if (Directory.Exists(curDir) == false)
            {
                Directory.CreateDirectory(curDir);
            }

            if (File.Exists(curDir + logFile) == false)
            {
                FileStream log = new FileStream(curDir + logFile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                log.Close();
                Console.WriteLine("Creating Log File.");
                StreamWriter writer = new StreamWriter(curDir + logFile);
                writer.WriteLine(@"*******************************************************************************");
                writer.Close();
                writer.Dispose();
            }

        }
    }
}
