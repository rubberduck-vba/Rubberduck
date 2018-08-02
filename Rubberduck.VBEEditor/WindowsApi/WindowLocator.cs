using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.WindowsApi
{
    public static class WindowLocator
    {
        public static IntPtr FindChildWindow(this IntPtr parentHandle, string caption)
        {
            return new ChildWindowFinder(parentHandle, caption).ResultHandle;
        }

        private class ChildWindowFinder
        {
            private readonly IntPtr _parentHandle;
            private readonly string _caption;

            public ChildWindowFinder(IntPtr parentHandle, string caption)
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
