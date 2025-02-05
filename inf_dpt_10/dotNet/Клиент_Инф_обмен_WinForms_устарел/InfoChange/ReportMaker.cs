using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;

namespace InfoChange
{
    class ReportMaker
    {
        public string path;
        public bool Started;
        public bool Ended;

        public ReportMaker()
        {
            ClearReport();
        }

        public bool isOOoInstalled()
        {
            try
            {
                string baseKey;
                //if ()

                baseKey = @"SOFTWARE\OpenOffice.org\";

                // Get the URE directory
                string key = baseKey + @"Layers\URE\1";
                RegistryKey reg = Registry.CurrentUser.OpenSubKey(key);
                if (reg == null) reg = Registry.LocalMachine.OpenSubKey(key);
                string urePath = reg.GetValue("UREINSTALLLOCATION") as string;
                reg.Close();
                if (urePath != null) return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }

        public bool isMSWordInstalled()
        {
            // Check whether Microsoft Word is installed on this computer,
            // by searching the HKEY_CLASSES_ROOT\Word.Application key.
            using (RegistryKey regWord = Registry.ClassesRoot.OpenSubKey("Word.Application"))
            {
                if (regWord == null)
                {
                    //Console.WriteLine("Microsoft Word is not installed");
                    return false;
                }
                else
                {
                    //Console.WriteLine("Microsoft Word is installed");
                    return true;
                }
            }
        }

        public string GetApplicationForOpen()
        {
            string txtAppName = "explorer";

            if (isOOoInstalled())
            {
                txtAppName = "swriter";
            }
            else if (isMSWordInstalled())
            {
                txtAppName = "winword";
            }

            return txtAppName;
        }

        public string getTempFolderPath()
        {
            string temp = Path.GetTempPath();
            return temp;
        }

        public string getRandFileName()
        {
            string file_name = Guid.NewGuid().ToString() + ".html"; ;
            return file_name;
        }

        private static bool WriteTofile(string txtText, string outfile)
        {
            using (StreamWriter sw = new StreamWriter(outfile, true))
            {
                sw.WriteLine(txtText);
                sw.Close();
            }
            return true;

        }
        public bool ClearReport()
        {
            Started = false;
            Ended = false;
            path = "";

            return true;
        }

        public bool StartReport()
        {
            try
            {
                if (!Started || Ended)
                {
                    // создать временный файл, записать туда отчет, вернуть ссылку на файл
                    path = Path.GetTempFileName();
                    // меняем раширение .tmp на .html
                    path = path.Substring(0, path.Length - 4) + ".html";

                    WriteTofile("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\">\n<HTML>\n\n<HEAD>\n   <META HTTP-EQUIV=\"CONTENT-TYPE\" CONTENT=\"text/html; charset=UTF-8\">\n	<TITLE></TITLE>\n</HEAD>\n<BODY style=\"font-size:8pt; font-family: Times;\">", path);
                    Started = true;
                    Ended = false;
                }
                else
                {
                    throw new System.ArgumentException("Report already started and not ended", "original");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }

            return true;
        }

        public bool EndReport()
        {
            try{
                if (Started && !Ended)
                {
                    if (File.Exists(path))
                    {
                        WriteTofile("</BODY>\n</HTML>", path);
                        Ended = true;
                    }
                    else
                    {
                        throw new System.ArgumentException("Destination report file not found", "original");
                    }
                }
                else
                {
                    if (!Started) { throw new System.ArgumentException("Report not started", "original");  }
                    else if (Ended) { throw new System.ArgumentException("Report already finished", "original"); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }

            return true;
        }

        public bool SplitNewPage()
        {
            
            return AddToReport(Convert.ToString("\f")); // \u000C = 12
        }

        public bool AddToReport(string txtNewReportString)
        {
            try
            {
                if (Started && !Ended)
                {
                    WriteTofile(txtNewReportString, path);
                }
                else
                {
                    if (!Started) { throw new System.ArgumentException("Report not started", "original"); }
                    else if (Ended) { throw new System.ArgumentException("Report already finished", "original"); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }
            return true;

        }

        public bool ShowReport()
        {
            try
            {
                Process proc = new Process();
                proc.StartInfo.FileName = GetApplicationForOpen();
                proc.StartInfo.Arguments = @"""" + path + @"""";
                //proc.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }

            return true;
        }


        public bool makeFakeReport(string htmlBody)
        {
            try
            {
                // создать временный файл, записать туда отчет, вернуть ссылку на файл
                string path = Path.GetTempFileName();
                // меняем раширение .tmp на .html
                path = path.Substring(0, path.Length - 4) + ".html";

                WriteTofile("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\">\n<HTML>\n\n<HEAD>\n   <META HTTP-EQUIV=\"CONTENT-TYPE\" CONTENT=\"text/html; charset=UTF-8\">\n	<TITLE>Call Center</TITLE>\n</HEAD>\n<BODY>", path);
                WriteTofile(htmlBody, path);
                WriteTofile("</BODY>\n</HTML>", path);

                Process proc = new Process();
                proc.StartInfo.FileName = GetApplicationForOpen();
                proc.StartInfo.Arguments = @"""" + path + @"""";
                //proc.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                proc.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка приложения. Message: " + ex.ToString(), "Внимание!", MessageBoxButtons.OK);
                return false;
            }

            return true;
                
        }

        ~ReportMaker()
        {
            try{
                if (path != null)
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
            }catch{
                ;
                // ну, значит кто-то держит файл - пусть держит дальше, удалить не выходит..
            }

        }



    }
}
