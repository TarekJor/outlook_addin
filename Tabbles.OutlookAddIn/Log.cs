using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tabbles.OutlookAddIn
{
    public static class Log
    {
        private static string getLogFilePath()
        {
            var folderDocs = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            var tabblesFolder = System.IO.Path.Combine(folderDocs, "Tabbles");
            System.IO.Directory.CreateDirectory(tabblesFolder);
            return (System.IO.Path.Combine(tabblesFolder, "log_outlook_addin.txt"));
        }

        /// <summary>
        /// Meant to be called every 10 minutes or so.
        /// </summary>
        public static void deleteLogIfTooLong()
        {

            var logFilePath = getLogFilePath();

            var fi = new System.IO.FileInfo(logFilePath);
            if (fi.Length > 1000000)
            {
                System.IO.File.Delete(logFilePath);
            }


        }

        public static void log(string txt)
        {
            try
            {

                var logFilePath = getLogFilePath();

                using (var sw = System.IO.File.AppendText(logFilePath))
                {
                    sw.WriteLine(DateTime.Now.ToString() + ":  " + txt + System.Environment.NewLine + System.Environment.NewLine);
                }
            }
            catch { }
        }

    }
}
