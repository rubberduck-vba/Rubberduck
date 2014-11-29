using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public static class NativeWindowMethods
    {
        [DllImport("user32", EntryPoint = "SendMessageW", ExactSpelling = true)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

        //todo: fix the delegate...
        public delegate int CallBackEnumWindowsDelegate(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int EnumChildWindows(IntPtr parentWindowHandle, CallBackEnumWindowsDelegate lpEnumFunction, IntPtr lParam);

        [DllImport("user32", EntryPoint = "GetWindowTextW", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("User32.dll")]
        static extern IntPtr GetParent(IntPtr hWnd);

        internal static string GetWindowTextByHwnd(IntPtr windowHandle)
        {
            const int MAX_BUFFER = 300;

            StringBuilder bufferStringBuilder = null;
            int charactersCount = 0;
            string result = null;

            bufferStringBuilder = new StringBuilder(MAX_BUFFER + 1);

            charactersCount = GetWindowText(windowHandle, bufferStringBuilder, MAX_BUFFER);
            if (charactersCount > 0)
            {
                result = bufferStringBuilder.ToString(0, charactersCount);
            }

            return result;
        }

        internal static void ActivateWindow(IntPtr windowHandle, IntPtr parentWindowHandle)
        {
            const int WM_MOUSEACTIVATE = 0x21;
            const int HTCAPTION = 2;
            const int WM_LBUTTONDOWN = 0x201;

            SendMessage(windowHandle, WM_MOUSEACTIVATE, parentWindowHandle, new IntPtr(HTCAPTION + WM_LBUTTONDOWN * 0x10000));
        }

        internal static void EnumChildWindows(IntPtr parentWindowHandle, CallBackEnumWindowsDelegate callBackEnumWindows)
        {
            int result;

            result = EnumChildWindows(parentWindowHandle, callBackEnumWindows, IntPtr.Zero);

            if (result != 0)
            {
                System.Diagnostics.Debug.WriteLine("EnumChildWindows failed");
            }
        }

    }

    internal class ChildWindowFinder
    {
        private IntPtr m_resultHandle = IntPtr.Zero;
        private string m_caption;

        internal ChildWindowFinder(string caption)
        {
            m_caption = caption;
        }

        public int EnumWindowsProcToChildWindowByCaption(IntPtr windowHandle, IntPtr param)
        {
            string caption;
            int result;

            // By default it will continue enumeration after this call
            result = 1;

            caption = NativeWindowMethods.GetWindowTextByHwnd(windowHandle);


            if (m_caption == caption)
            {
                // Found
                m_resultHandle = windowHandle;

                // Stop enumeration after this call
                result = 0;
            }
            return result;
        }

        public IntPtr ResultHandle
        {
            get
            {
                return m_resultHandle;
            }
        }
    }
}
