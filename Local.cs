using ClosedXML.Excel;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using ekmui_tester.FW.tools.prvt;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ekmui_tester.FW.tools
{
    public static class Local
    {
        #region  system paths
        public static string Automation_dir = @"C:\automation\";
        public static string Temp_folder = @"C:\automation\temp\";
        public static string Data_folder = @"C:\automation\data\";
        public static string Log_Path = @"C:\automation\logs\";        
        public static string Counter_Path = Temp_folder+"counter.dat";
        public static string Record_File = Log_Path + "bug_record";        
        public static string output_file = Log_Path + "output.log";


        public static string EKM_data_path = @"C:\ProgramData\Dyadic\ekm\";

        #endregion

        //redirect cmd to file
        private static bool cmd_redirected = false;
        public static StreamWriter stream;
        public static FileStream fs;
        public static TextWriter originalCMD;


        #region files and directories

        public static bool folderExists(string path)
        {
            return Directory.Exists(path);
        }

        public static bool fileExists(string filepath)
        {            
            return File.Exists(filepath);
            
        }

        public static void WriteToExcel(DataTable table, string file_name)
        {
            try
            {
                XLWorkbook wk = new XLWorkbook();
                DataSet set = new DataSet();
                set.Tables.Add(table);
                wk.AddWorksheet(set);
                wk.SaveAs(file_name);
                Logger.success("excel succesfully saved at " + file_name);
            }
            catch (Exception e)
            {
                Logger.failure("error saving excel file");
                Console.WriteLine(e);
            }
        }

        /// <summary>
        /// moves a file locally to EKM_data_path
        /// </summary>
        /// <param name="origin">file orgin path</param>
        /// <param name="replace">replace file if already exists</param>
        /// <returns>if action sucssed</returns>        
        public static bool MoveFile(string origin, string destiniation, bool replace)
        {
            try
            {
                if (replace)
                    if (fileExists(destiniation))
                        File.Delete(destiniation);

                File.Move(origin, destiniation);
                Logger.info("File " + origin + " moved to " + destiniation);
            }
            catch (Exception e)
            {
                Logger.failure("Error moving file '" + origin + "'. exception: ");
                Console.WriteLine(e + "\n\n");
                return false;
            }

            return true;
        }

        internal static void Save_txtFile(string path, string text)
        {
            try
            {
                File.WriteAllText(path, text);
            }
            catch (Exception e)
            {

                Logger.failure("Error writing file to " + path);
                Console.WriteLine(e + "\n\n");
                
            }
           
        }

        internal static string Read_txtFile(string p)
        {
           try
            {
                return File.ReadAllText(p);              
            }
            catch (Exception e)
            {

                Logger.failure("Error reading file from " + p);
                Console.WriteLine(e + "\n\n");
            }
            return "";
        }

        #endregion

        #region cmd

        /// <summary>
        /// runs a command locally
        /// </summary>    
        /// <returns>if execuation fails returns 'X', else returns result</returns>
        public static string run(string command)
        {
            try
            {
                return Cmdrun(command);
            }
            catch (Exception e)
            {
                ekmui_tester.Logger.failure("cmd run failed. command: " + command);
                ekmui_tester.Logger.failure("error: " + e.Message);
                return "X";
            }

        }


        private static string Cmdrun(string command)
        {
            var process = new Process()
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = @"c:\Windows\System32\cmd.exe",
                    Arguments = $"/c \"{command}\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                }
            };
            process.Start();
            string result = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            return result;
        }


        internal static void close_output()
        {
            if (cmd_redirected)
            {
                stream.Flush();
                stream.Close();
            }
        }

        public static void RedirectCMD()
        {
            if (!cmd_redirected)
            {
                fs = new FileStream(output_file, FileMode.Create);
                stream = new StreamWriter(fs);
                // save the original standard output.                    
                originalCMD = Console.Out;
                Console.SetOut(stream);
                cmd_redirected = true;
            }
        }

        public static void stopRedirectCMD()
        {
            if (cmd_redirected)
            {
                Console.SetOut(originalCMD);
                cmd_redirected = false;
            }
        }

        #endregion

        #region browser
        #endregion

        #region chrome
        public static IntPtr isOpenShowed(TimeSpan time)
        {
            IntPtr open_window = IntPtr.Zero;
            int mil = 0;
            while (open_window == IntPtr.Zero && mil < time.TotalMilliseconds)
            {
                open_window = cWin32Api.FindWindow(null, "Choose File to Upload"); //IE
                if (open_window != IntPtr.Zero)
                    break;
                open_window = cWin32Api.FindWindow(null, "Save As"); //Chrome
                if (open_window != IntPtr.Zero)
                    break;

                open_window = cWin32Api.FindWindow(null, "Open"); //Chrome
                if (open_window != IntPtr.Zero)
                    break;

                System.Threading.Thread.Sleep(1000);
                mil += 1000;
            }

            if (mil >= time.TotalMilliseconds)
                return IntPtr.Zero;

            return open_window;

        }

        public static string Chrome_openWin(string file_path, IntPtr open_window, bool counter_extantion)
        {
            string extantion = "";
            if (counter_extantion)
                extantion = "_" + autotest.Counter;

            if (open_window != IntPtr.Zero)
            {

                //Get the open button and file path textbox HWND
                IntPtr OK_HWndControl = cWin32Api.GetDlgItem(open_window, 1);

                StringBuilder class_name = new StringBuilder(256);

                IntPtr Path_HWndControl =
                  cWin32Api.GetAllChildHandles(
                            cWin32Api.GetAllChildHandles(open_window).Find(x =>
                            {
                                cWin32Api.GetClassName(x, class_name, 256);
                                return class_name.ToString().Equals("ComboBox");
                            })
                  ).Find(x =>
                  {
                      cWin32Api.GetClassName(x, class_name, 256);
                      return class_name.ToString().Equals("Edit");
                  });


                HandleRef hrefHWndTarget = new HandleRef(null, Path_HWndControl);

                //Send message with the file path
                string full_path = file_path + extantion;
                System.Threading.Thread.Sleep(500);
                cWin32Api.SendMessage(hrefHWndTarget, (uint)cWin32Api.MessageValue.WM_SETTEXT, 0, full_path);
                Logger.info("window_handle path: " + file_path + extantion);
                System.Threading.Thread.Sleep(500);
                //  cWin32Api.SendMessage(Path_HWndControl, (uint)cWin32Api.MessageValue.WM_SETTEXT,0, "hi");
                //Send message to click on Open
                cWin32Api.SendMessage(OK_HWndControl, (int)cWin32Api.MessageValue.BN_CLICKED, 0, IntPtr.Zero);
                Logger.info("closed chrome open-window");
                return extantion;
            }

            Logger.info("chrome open-window never showed");
            return "";
        }

        public static string Chrome_Upload(string file_path)
        {
            Logger.info("OpenDLG ->");
            System.Threading.Thread.Sleep(300);

            //TestStack.White.Application app = TestStack.White.Application.Attach("chrome.exe");
            //Window win = app.GetWindow("Save As");
            //Button button = win.Get<Button>(SearchCriteria.ByText("Save"));

            //Get the open file dialog handle
            string file = Chrome_openWin(Data_folder + file_path, isOpenShowed(TimeSpan.FromSeconds(15)), false);
            if (file == "")
                return null;

            return file_path + file;

        }

        public static string Chrome_Download(string file_path)
        {
            Logger.info("OpenDLG ->");
            System.Threading.Thread.Sleep(300);

            //TestStack.White.Application app = TestStack.White.Application.Attach("chrome.exe");
            //Window win = app.GetWindow("Save As");
            //Button button = win.Get<Button>(SearchCriteria.ByText("Save"));

            //Get the open file dialog handle            
            string file = Chrome_openWin(Temp_folder + file_path, isOpenShowed(TimeSpan.FromSeconds(15)), true);

            if (file == "")
                return null;

            return file_path + file;

        }
        #endregion

        #region firefox
        #endregion

    }
}
