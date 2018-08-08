using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.WindowsApi
{
    public static class WindowLocator
    {
        public static bool IsChildOf(this IntPtr childHandle, IntPtr hWnd)
        {
            return hWnd.Equals(User32.GetAncestor(childHandle, User32.GetAncestorFlags.GetRoot));
        }

        public static List<IntPtr> ChildWindows(this IntPtr hWnd)
        {
            var children = new List<IntPtr>();
            var childAfter = IntPtr.Zero;
            while (true)
            {
                var located = User32.FindWindowEx(hWnd, childAfter, null, null);
                if (located == IntPtr.Zero)
                {
                    break;
                }
                children.Add(located);
                childAfter = located;
            }
            return children;
        }

        public static IntPtr FindChildWindow(this IntPtr parentHandle, string caption)
        {
            return new ChildWindowCaptionFinder(parentHandle, caption).ResultHandle;
        }

        private class ChildWindowCaptionFinder
        {
            private readonly IntPtr _parentHandle;
            private readonly string _caption;

            public ChildWindowCaptionFinder(IntPtr parentHandle, string caption)
            {
                _parentHandle = parentHandle;
                _caption = caption;
            }

            private int EnumWindowsProcToChildWindowByCaption(IntPtr windowHandle, IntPtr param)
            {
                // By default it will continue enumeration after this call
                var result = 1;
                var caption = windowHandle.GetWindowText();

                if (_caption == caption)
                {
                    // Found
                    _resultHandle = windowHandle;

                    // Stop enumeration after this call
                    result = 0;
                }
                return result;
            }

            private IntPtr _resultHandle = IntPtr.Zero;

            public IntPtr ResultHandle
            {
                get
                {
                    NativeMethods.EnumChildWindows(_parentHandle, EnumWindowsProcToChildWindowByCaption);
                    return _resultHandle;
                }
            }
        }
    }
}
