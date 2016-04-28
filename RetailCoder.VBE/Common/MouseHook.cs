using Rubberduck.Common.WinAPI;
using Rubberduck.UI.Command.MenuItems;
using System;

namespace Rubberduck.Common
{
    public sealed class MouseHook : LowLevelHook
    {
        private readonly IntPtr _vbeHandle;

        public MouseHook(IntPtr vbeHandle) : base(WindowsHook.MOUSE_LL)
        {
            _vbeHandle = vbeHandle;
        }

        protected override void HookCallbackCore(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (
                ((WM)wParam == WM.RBUTTONUP || (WM)wParam == WM.LBUTTONUP)
                && User32.IsVbeWindowActive(_vbeHandle))
            {
                UiDispatcher.InvokeAsync(() => OnMessageReceived());
            }
        }
    }
}