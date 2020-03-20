using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ekmui_tester.FW.tools.prvt;

namespace ekmui_tester.FW.tools
{
    class WindowHandle 
    {                
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

                System.Threading.Thread.Sleep(1000);
                mil+=1000;
            }

            if (mil >= time.TotalMilliseconds)
                return IntPtr.Zero;

            return open_window;

        }

        public static string Chrome_openWin(string file_path, IntPtr open_window)
        {
            string extantion = "_" + autotest.Counter;

            if (open_window != IntPtr.Zero)
            {
                List<IntPtr> children = cWin32Api.GetAllChildHandles(open_window);

                //Get the open button and file path textbox HWND
                IntPtr OK_HWndControl = cWin32Api.GetDlgItem(open_window, 1);
                int a = 1148;
                IntPtr Path_HWndControl = children[4];
                HandleRef hrefHWndTarget = new HandleRef(null, Path_HWndControl);

                //Send message with the file path
                string full_path = file_path + extantion;
                cWin32Api.SendMessage(hrefHWndTarget, (uint)cWin32Api.MessageValue.WM_SETTEXT, 0, full_path);
                Logger.info("window_handle path: " + file_path + extantion);
                System.Threading.Thread.Sleep(500);
                //  cWin32Api.SendMessage(Path_HWndControl, (uint)cWin32Api.MessageValue.WM_SETTEXT,0, "hi");
                //Send message to click on Open
                cWin32Api.SendMessage(OK_HWndControl, (int)cWin32Api.MessageValue.BN_CLICKED, 0, IntPtr.Zero);
                autotest.Counter++;
                Logger.info("closed chrome open-window");
                return full_path;
            }
            
            Logger.info("chrome open-window never showed");
            return "";                
        }
        
        public static string Chrome_openWin(string file_path)
        {            
            Logger.info("OpenDLG ->");
            System.Threading.Thread.Sleep(300);               

            //TestStack.White.Application app = TestStack.White.Application.Attach("chrome.exe");
            //Window win = app.GetWindow("Save As");
            //Button button = win.Get<Button>(SearchCriteria.ByText("Save"));

            //Get the open file dialog handle
          return Chrome_openWin(file_path, isOpenShowed(TimeSpan.FromSeconds(15)));

        }
      
    }
}
