﻿using System;

namespace Rubberduck.VBEditor.WindowsApi
{
    internal class ChildWindowFinder
    {
        private readonly string _caption;

        internal ChildWindowFinder(string caption)
        {
            _caption = caption;
        }

        public int EnumWindowsProcToChildWindowByCaption(IntPtr windowHandle, IntPtr param)
        {
            // By default it will continue enumeration after this call
            var result = 1;
            var caption = windowHandle.GetWindowText();

            if (_caption == caption)
            {
                // Found
                ResultHandle = windowHandle;

                // Stop enumeration after this call
                result = 0;
            }
            return result;
        }

        public IntPtr ResultHandle { get; private set; } = IntPtr.Zero;
    }
}
