using Rubberduck.Common.WinAPI;
using Rubberduck.UI.Command.MenuItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Rubberduck.Common
{
    public sealed class KeyboardHook : LowLevelHook
    {
        // keys that don't modify anything in the code pane but the current selection
        private static readonly IReadOnlyList<Keys> NavigationKeys = new[]
        {
            Keys.Down,
            Keys.Up,
            Keys.Left,
            Keys.Right,
            Keys.PageDown,
            Keys.PageUp,
            Keys.Home,
            Keys.End,
        };

        private readonly IntPtr _vbeHandle;

        public KeyboardHook(IntPtr vbeHandle) : base(WindowsHook.KEYBOARD_LL)
        {
            _vbeHandle = vbeHandle;
        }

        protected override void HookCallbackCore(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (
                (WM)wParam == WM.KEYUP
                && NavigationKeys.Contains((Keys)Marshal.ReadInt32(lParam))
                && User32.IsVbeWindowActive(_vbeHandle))
            {
                UiDispatcher.InvokeAsync(() => OnMessageReceived());
            }
        }
    }
}