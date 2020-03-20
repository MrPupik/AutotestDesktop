using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text;

namespace ekmui_tester.FW.tools.prvt
{
    internal static class cWin32Api
    {

        #region ENUM
        [Flags]
        [TypeConverter(typeof(Int32))]
        public enum MessageValue
        {
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205,
            MN_GETHMENU = 0x1e1,
            WM_KEYDOWN = 0x100,
            WM_KEYUP = 0x101,
            WM_ENTER = 7001,
            WM_COMMAND = 0x111,
            VK_NUMLOCK = 0x90,
            WM_USER = 0x400,
            WM_AUTO_TEST_MSG = WM_USER + 169,
            MK_RBUTTON = 0x0002,
            SW_SHOW = 5,
            MAXTITLE = 255,
            VK_ESCAPE = 0x01b,
            VK_TAB = 0x09,
            VK_SPACE = 0x20,
            WM_CLOSE = 16,
            BN_CLICKED = 245,
            WM_SETTEXT = 0x000C,
        };
        #endregion

        #region Variables
        private const int SW_SHOW = 5;
        const int MAXTITLE = 255;
        private static ArrayList mTitlesList;
        private delegate bool EnumDelegate(IntPtr hWnd, int lParam);
        #endregion

        #region Win32 Api Functions
        [DllImportAttribute("user32.dll")]
        private static extern bool EnumDesktopWindows(IntPtr hDesktop,
                                                       EnumDelegate lpEnumCallbackFunction,
                                                       IntPtr lParam);
        [DllImportAttribute("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd,
                                                 StringBuilder lpWindowText,
                                                 int nMaxCount);
       

        public delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);


        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

       
    

        [DllImport("user32.dll", SetLastError = false)]
        public static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

        [DllImport("kernel32.dll")]
        public static extern bool GetVersionEx(ref OSVERSIONINFO lpVersionInfo);

        //[DllImportAttribute("User32.dll")]
        //internal static extern IntPtr FindWindow(string ClassName, string WindowName);

        [DllImportAttribute("User32.dll")]
        private static extern IntPtr SetForegroundWindow(IntPtr hWnd);

        [DllImportAttribute("User32.dll")]
        internal static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hWnd, int msg, int wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        public static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, int wParam, string lParam);



        [DllImport("user32.dll", EntryPoint = "GetClassName")]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);

        [DllImportAttribute("user32.dll")]
        public static extern IntPtr GetMenu(IntPtr hWnd);

        [DllImportAttribute("user32.dll")]
        public static extern int GetMenuItemCount(IntPtr hWnd);

        [DllImportAttribute("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool CloseWindow(IntPtr hWnd);


        [DllImport("advapi32.DLL", SetLastError = true)]
        public static extern int LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        #endregion
        #region Structers
        [StructLayout(LayoutKind.Sequential)]
        public struct OSVERSIONINFO
        {
            public System.Int32 dwOSVersionInfoSize;
            public System.Int32 dwMajorVersion;
            public System.Int32 dwMinorVersion;
            public System.Int32 dwBuildNumber;
            public System.Int32 dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public String szCSDVersion;
        }
        #endregion



        #region child_windows

        public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr lParam);





        //public static WindowHandleInfo(IntPtr handle)
        //{
        //    _MainHandle = handle;
        //}
        public static IntPtr _MainHandle;

        public static List<IntPtr> GetAllChildHandles(IntPtr _mainHandle)
        {
            _MainHandle = _mainHandle;
            List<IntPtr> childHandles = new List<IntPtr>();

            GCHandle gcChildhandlesList = GCHandle.Alloc(childHandles);
            IntPtr pointerChildHandlesList = GCHandle.ToIntPtr(gcChildhandlesList);

            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(_MainHandle, childProc, pointerChildHandlesList);
            }
            finally
            {
                gcChildhandlesList.Free();
            }

            return childHandles;
        }

        private static bool EnumWindow(IntPtr hWnd, IntPtr lParam)
        {
            GCHandle gcChildhandlesList = GCHandle.FromIntPtr(lParam);

            if (gcChildhandlesList == null || gcChildhandlesList.Target == null)
            {
                return false;
            }

            List<IntPtr> childHandles = gcChildhandlesList.Target as List<IntPtr>;
            childHandles.Add(hWnd);

            return true;
        }

        #endregion











        #region Function
        /// <summary>
        /// Constructor
        /// </summary>
        static cWin32Api()
        {
        }

        /// <summary>
        /// Show and hide consol
        /// </summary>
        /// <param name="Status"></param>
        /// <param name="Message">Messega to print on the console</param>
        public static void DisplayConsole(bool Status, string Window_Title)
        {
            IntPtr HWND = cWin32Api.FindWindow(null, Window_Title);

            if (Status)//Show
                cWin32Api.ShowWindow(HWND, 1);
            else//Hide
                cWin32Api.ShowWindow(HWND, 0);
        }

        /// <summary>
        /// Focuse a window
        /// </summary>
        /// <param name="WindowTitle">Window Title</param>
        public static void SetFocuse(string WindowTitle)
        {
            try
            {
                string[] winTitle = GetDesktopWindowsCaptions();
                for (int i = 0; i < winTitle.Length; i++)
                {
                    if (winTitle[i].Contains(WindowTitle) == true)
                    {
                        WindowTitle = winTitle[i];
                        break;
                    }
                }

                IntPtr HWND = FindWindow(null, WindowTitle);
                if (HWND != null)
                {
                    SetForegroundWindow(HWND);
                    ShowWindow(HWND, SW_SHOW);
                }
            }
            catch //(Exception ex) 
            {
                // MessageBox.Show(ex.Message); 
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private static bool EnumWindowsProc(IntPtr hWnd, int lParam)
        {
            string title = GetWindowText(hWnd);
            if (title != "")
                mTitlesList.Add(title);
            return true;
        }

        /// <summary>
        /// Returns the caption of a windows by given HWND identifier.
        /// </summary>
        public static string GetWindowText(IntPtr hWnd)
        {
            StringBuilder title = new StringBuilder(MAXTITLE);
            int titleLength = GetWindowText(hWnd, title, title.Capacity + 1);
            title.Length = titleLength;

            return title.ToString();
        }

        /// <summary>
        /// Returns the caption of all desktop windows.
        /// </summary>
        public static string[] GetDesktopWindowsCaptions()
        {
            mTitlesList = new ArrayList();
            EnumDelegate enumfunc = new EnumDelegate(EnumWindowsProc);
            IntPtr hDesktop = IntPtr.Zero; // current desktop
            bool success = EnumDesktopWindows(hDesktop, enumfunc, IntPtr.Zero);

            if (success)
            {
                // Copy the result to string array
                string[] titles = new string[mTitlesList.Count];
                mTitlesList.CopyTo(titles);
                return titles;
            }
            else
            {
                // Get the last Win32 error code
                int errorCode = Marshal.GetLastWin32Error();

                string errorMessage = String.Format(
                "EnumDesktopWindows failed with code {0}.", errorCode);
                throw new Exception(errorMessage);
            }
        }

        /// <summary>
        /// Show or Hide window
        /// 0 = SW_HIDE
        /// 1 = SW_SHOW
        /// </summary>
        /// <param name="ShowHide">Show or Hide</param>
        /// <param name="Title">Window Title</param>
        public static void DisplayWindow(int ShowHide, string Title)
        {
            IntPtr hWnd = FindWindow(null, Title);
            if (hWnd != IntPtr.Zero)
                ShowWindow(hWnd, ShowHide);
        }

        /// <summary>
        /// Close popup message 
        /// </summary>
        /// <param name="Title"></param>
        public static void Close_Popup_Message_By_Title(string Title)
        {
            try
            {
                IntPtr HWND = cWin32Api.FindWindow(null, Title);
                if (HWND != IntPtr.Zero)
                {
                    Console.WriteLine("HWND: " + HWND.ToString());
                    PostMessage(HWND, (int)cWin32Api.MessageValue.WM_KEYDOWN, (int)cWin32Api.MessageValue.VK_ESCAPE, IntPtr.Zero);
                    PostMessage(HWND, (int)cWin32Api.MessageValue.WM_KEYUP, (int)cWin32Api.MessageValue.VK_ESCAPE, IntPtr.Zero);
                    HWND = IntPtr.Zero;
                }
            }
            catch { }
        }
        #endregion
    }
}