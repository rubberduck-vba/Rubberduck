using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Rubberduck.VBEditor.WindowsApi
{
    public static class User32
    {
        #region WinEvents

        //https://msdn.microsoft.com/en-us/library/windows/desktop/dd373885(v=vs.85).aspx
        public delegate void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        //https://msdn.microsoft.com/en-us/library/windows/desktop/dd373640(v=vs.85).aspx
        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventProc lpfnWinEventProc, uint idProcess, uint idThread, WinEventFlags dwFlags);

        /// <summary>
        /// Removes event hooks set with SetWinEventHook.
        /// https://msdn.microsoft.com/en-us/library/windows/desktop/dd373671(v=vs.85).aspx
        /// </summary>
        /// <param name="hWinEventHook">The hook handle to unregister.</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        #endregion

        /// <summary>
        /// Returns the thread ID for thread that created the passed hWnd.
        /// https://msdn.microsoft.com/en-us/library/windows/desktop/ms633522(v=vs.85).aspx
        /// </summary>
        /// <param name="hWnd">The window handle to get the thread ID for.</param>
        /// <param name="processId">This is actually an out parameter in the API, but we don't care about it. Should always be IntPtr.Zero.</param>
        /// <returns>Unmanaged thread ID</returns>
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

        /// <summary>
        /// Retrieves the identifier of the thread that created the specified window and, optionally, 
        /// the identifier of the process that created the window.
        /// </summary>
        /// <param name="hWnd">A handle to the window.</param>
        /// <param name="lpdwProcessId">A pointer to a variable that receives the process identifier. 
        /// If this parameter is not NULL, GetWindowThreadProcessId copies the identifier of the process to the variable; otherwise, it does not.</param>
        /// <returns>The return value is the identifier of the thread that created the window.</returns>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        /// <summary>
        /// Retrieves a handle to the foreground window (the window with which the user is currently working). 
        /// The system assigns a slightly higher priority to the thread that creates the foreground window than it does to other threads.
        /// </summary>
        /// <returns>The return value is a handle to the foreground window. 
        /// The foreground window can be NULL in certain circumstances, such as when a window is losing activation.</returns>
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr GetActiveWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr GetFocus();

        /// <summary>
        /// Gets the underlying class name for a window handle.
        /// https://msdn.microsoft.com/en-us/library/windows/desktop/ms633582(v=vs.85).aspx
        /// </summary>
        /// <param name="hWnd">The handle to retrieve the name for.</param>
        /// <param name="lpClassName">Buffer for returning the class name.</param>
        /// <param name="nMaxCount">Buffer size in characters, including the null terminator.</param>
        /// <returns>The length of the returned class name (without the null terminator), zero on error.</returns>
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        /// <summary>   Gets the parent window of this item. </summary>
        ///
        /// <param name="hWnd"> The window handle. </param>
        /// <returns>   The parent window IntPtr handle. </returns>
        [DllImport("User32.dll")]
        internal static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string lclassName, string windowTitle);

        /// <summary>
        /// Validates a window handle.
        /// </summary>
        /// <param name="hWnd">The handle to validate.</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindow(IntPtr hWnd);
    }
}
