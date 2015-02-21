using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace Rubberduck.UI
{
    /// <summary>
    /// Collection of WinAPI methods and extensions to handle native windows.
    /// </summary>
    // Special Thank You to Carlos Quintero for supplying the project with the original code this file is based on.
    public static class NativeWindowMethods
    {
        /// <summary>   Sends a message to the OS. </summary>
        ///
        /// <param name="hWnd">     The window handle. </param>
        /// <param name="wMsg">     The message. </param>
        /// <param name="wParam">   The parameter. </param>
        /// <param name="lParam">   The parameter. </param>
        /// <returns>   An IntPtr handle. </returns>
        [DllImport("user32", EntryPoint = "SendMessageW", ExactSpelling = true)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

        /// <summary>   EnumChildWindows delegate. </summary>
        ///
        /// <param name="hwnd"> Main Window Handle</param>
        /// <param name="lParam"> Application defined parameter. Unused. </param>
        /// <returns>   An int. </returns>
        public delegate int EnumChildWindowsDelegate(IntPtr hwnd, IntPtr lParam);

        /// <summary>   WinAPI method to Enumerate Child Windows </summary>
        ///
        /// <param name="parentWindowHandle">   Handle of the parent window. </param>
        /// <param name="lpEnumFunction">       The enum delegate function. </param>
        /// <param name="lParam">               The parameter. </param>
        /// <returns>   An int. </returns>
        [DllImport("user32", ExactSpelling = true, CharSet = CharSet.Unicode)]
        public static extern int EnumChildWindows(IntPtr parentWindowHandle, EnumChildWindowsDelegate lpEnumFunction, IntPtr lParam);

        /// <summary>   Gets window text. </summary>
        ///
        /// <param name="hWnd">         The window handle. </param>
        /// <param name="lpString">     The return string. </param>
        /// <param name="nMaxCount">    Number of maximums. </param>
        /// <returns>   Integer Success Code </returns>
        [DllImport("user32", EntryPoint = "GetWindowTextW", ExactSpelling = true, CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        /// <summary>   Gets the parent window of this item. </summary>
        ///
        /// <param name="hWnd"> The window handle. </param>
        /// <returns>   The parent window IntPtr handle. </returns>
        [DllImport("User32.dll")]
        public static extern IntPtr GetParent(IntPtr hWnd);

        /// <summary>   Gets window caption text by handle. </summary>
        ///
        /// <param name="windowHandle"> Handle of the window to be activated. </param>
        /// <returns>   The window caption text. </returns>
        public static string GetWindowTextByHwnd(IntPtr windowHandle)
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

        /// <summary>Activates the window by simulating a click.</summary>
        ///
        /// <param name="windowHandle">         Handle of the window to be activated. </param>
        /// <param name="parentWindowHandle">   Handle of the parent window. </param>
        public static void ActivateWindow(IntPtr windowHandle, IntPtr parentWindowHandle)
        {
            const int WM_MOUSEACTIVATE = 0x21;
            const int HTCAPTION = 2;
            const int WM_LBUTTONDOWN = 0x201;

            SendMessage(windowHandle, WM_MOUSEACTIVATE, parentWindowHandle, new IntPtr(HTCAPTION + WM_LBUTTONDOWN * 0x10000));
        }

        internal static void EnumChildWindows(IntPtr parentWindowHandle, EnumChildWindowsDelegate callBackEnumWindows)
        {
            int result;

            result = EnumChildWindows(parentWindowHandle, callBackEnumWindows, IntPtr.Zero);

            if (result != 0)
            {
                System.Diagnostics.Debug.WriteLine("EnumChildWindows failed");
            }
        }

        internal class ChildWindowFinder
        {
            private IntPtr _resultHandle = IntPtr.Zero;
            private string _caption;

            internal ChildWindowFinder(string caption)
            {
                _caption = caption;
            }

            public int EnumWindowsProcToChildWindowByCaption(IntPtr windowHandle, IntPtr param)
            {
                string caption;
                int result;

                // By default it will continue enumeration after this call
                result = 1;

                caption = NativeWindowMethods.GetWindowTextByHwnd(windowHandle);


                if (_caption == caption)
                {
                    // Found
                    _resultHandle = windowHandle;

                    // Stop enumeration after this call
                    result = 0;
                }
                return result;
            }

            public IntPtr ResultHandle
            {
                get
                {
                    return _resultHandle;
                }
            }
        }
    }
}
