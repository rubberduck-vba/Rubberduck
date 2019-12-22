using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Rubberduck.VBEditor.WindowsApi
{
    /// <summary>
    /// Collection of WinAPI methods and extensions to handle native windows.
    /// </summary>
    /// <remarks>
    /// **Special Thanks** to Carlos Quintero for supplying the project with the original code this file is based on.
    /// </remarks>
    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll")]
        public static extern bool TrackPopupMenuEx(IntPtr hmenu, uint fuFlags, int x, int y, IntPtr hwnd, IntPtr lptpm);

        [DllImport("user32.dll", EntryPoint = "InsertMenuW", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool InsertMenu(IntPtr hMenu, uint wPosition, uint wFlags, UIntPtr wIDNewItem, [MarshalAs(UnmanagedType.LPWStr)]string lpNewItem);

        [DllImport("user32.dll")]
        public static extern bool DestroyMenu(IntPtr hMenu);

        /// <summary>   Sends a message to the OS. </summary>
        ///
        /// <param name="hWnd">     The window handle. </param>
        /// <param name="wMsg">     The message. </param>
        /// <param name="wParam">   The parameter. </param>
        /// <param name="lParam">   The parameter. </param>
        /// <returns>   An IntPtr handle. </returns>
        [DllImport("user32", EntryPoint = "SendMessageW", ExactSpelling = true)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct NativeMessage
        {
            public IntPtr handle;
            public uint msg;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public System.Drawing.Point p;
        }

        [Flags]
        public enum PeekMessageRemoval : uint
        {
            NoRemove = 0,
            Remove = 1,
            NoYield = 2
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PeekMessage(ref NativeMessage lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, PeekMessageRemoval wRemoveMsg);

        [DllImport("user32.dll")]
        public static extern bool TranslateMessage([In] ref NativeMessage lpMsg);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage([In] ref NativeMessage lpMsg);

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

        /// <summary> Retrieves the name of the class to which the specified window belongs. </summary>
        /// <param name="hWnd">A handle to the window and, indirectly, the class to which the window belongs.</param>
        /// <param name="lpClassName">The class name string.</param>
        /// <param name="nMaxCount">The length of the <see cref="lpClassName" /> buffer, in characters. The buffer must be large enough to include the terminating null character; otherwise, the class name string is truncated to <see cref="nMaxCount" /> characters.</param>
        /// <returns>If the function succeeds, the return value is the number of characters copied to the buffer, not including the terminating null character. If the function fails, the return value is zero.To get extended error information, call GetLastError.</returns>
        [DllImport("user32.dll", EntryPoint = "GetClassNameW", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        /// <summary>   Gets window text. </summary>
        ///
        /// <param name="hWnd">         The window handle. </param>
        /// <param name="lpString">     The return string. </param>
        /// <param name="nMaxCount">    Number of maximums. </param>
        /// <returns>   Integer Success Code </returns>
        [DllImport("user32", EntryPoint = "GetWindowTextW", ExactSpelling = true, CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);


        /// <summary>   Gets the child window at the top of the Z order. </summary>
        ///
        /// <param name="hWnd"> The window handle. </param>
        /// <returns>   The child window IntPtr handle. </returns>
        [DllImport("user32.dll")]
        public static extern IntPtr GetTopWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string lclassName, string windowTitle);

        /// <summary>
        /// Forces the specified region to repaint.
        /// </summary>
        /// <param name="hWnd">The window to redraw</param>
        /// <param name="lprcUpdate">The update region. Ignored if hrgnUpdate is not null.</param>
        /// <param name="hrgnUpdate">Handle to an update region. Defaults to the entire window if null and lprcUpdate is null</param>
        /// <param name="flags">Redraw flags.</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, RedrawWindowFlags flags);

        /// <summary>   Gets window caption text by handle. </summary>
        ///
        /// <param name="windowHandle"> Handle of the window to be activated. </param>
        /// <returns>   The window caption text. </returns>
        public static string GetWindowText(this IntPtr windowHandle)
        {
            const int MAX_BUFFER = 300;

            var result = string.Empty;
            var bufferStringBuilder = new StringBuilder(MAX_BUFFER + 1);

            var charactersCount = GetWindowText(windowHandle, bufferStringBuilder, MAX_BUFFER);
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

        public static void EnumChildWindows(IntPtr parentWindowHandle, EnumChildWindowsDelegate callBackEnumWindows)
        {
            var result = EnumChildWindows(parentWindowHandle, callBackEnumWindows, IntPtr.Zero);
            if (result != 0)
            {
                Debug.WriteLine("EnumChildWindows failed");
            }
        }
    }

    [Flags]
    public enum RedrawWindowFlags : uint
    {
        /// <summary>
        /// Invalidates the rectangle or region that you specify in lprcUpdate or hrgnUpdate.
        /// You can set only one of these parameters to a non-NULL value. If both are NULL, RDW_INVALIDATE invalidates the entire window.
        /// </summary>
        Invalidate = 0x1,

        /// <summary>Causes the OS to post a WM_PAINT message to the window regardless of whether a portion of the window is invalid.</summary>
        InternalPaint = 0x2,

        /// <summary>
        /// Causes the window to receive a WM_ERASEBKGND message when the window is repainted.
        /// Specify this value in combination with the RDW_INVALIDATE value; otherwise, RDW_ERASE has no effect.
        /// </summary>
        Erase = 0x4,

        /// <summary>
        /// Validates the rectangle or region that you specify in lprcUpdate or hrgnUpdate.
        /// You can set only one of these parameters to a non-NULL value. If both are NULL, RDW_VALIDATE validates the entire window.
        /// This value does not affect internal WM_PAINT messages.
        /// </summary>
        Validate = 0x8,

        NoInternalPaint = 0x10,

        /// <summary>Suppresses any pending WM_ERASEBKGND messages.</summary>
        NoErase = 0x20,

        /// <summary>Excludes child windows, if any, from the repainting operation.</summary>
        NoChildren = 0x40,

        /// <summary>Includes child windows, if any, in the repainting operation.</summary>
        AllChildren = 0x80,

        /// <summary>
        /// Causes the affected windows, which you specify by setting the RDW_ALLCHILDREN and RDW_NOCHILDREN values, to receive
        /// WM_ERASEBKGND and WM_PAINT messages before the RedrawWindow returns, if necessary.
        /// </summary>
        UpdateNow = 0x100,

        /// <summary>
        /// Causes the affected windows, which you specify by setting the RDW_ALLCHILDREN and RDW_NOCHILDREN values, to receive WM_ERASEBKGND
        /// messages before RedrawWindow returns, if necessary.
        /// The affected windows receive WM_PAINT messages at the ordinary time.
        /// </summary>
        EraseNow = 0x200,

        Frame = 0x400,

        NoFrame = 0x800
    }
}
