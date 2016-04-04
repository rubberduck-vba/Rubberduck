using System;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class LowLevelMouseHook : IAttachable, IDisposable
    {
        private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && (WM)wParam == WM.RBUTTONDOWN)
            {
                OnRightClickCaptured();
            }
            return User32.CallNextHookEx(_hookId, code, wParam, lParam);
        }

        public event EventHandler RightClickCaptured;
        private void OnRightClickCaptured()
        {
            var handler = RightClickCaptured;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        //private const int WH_MOUSE_LL = 14;

        private IntPtr _hookId = IntPtr.Zero;

        public bool IsAttached { get; private set; }
        public event EventHandler<HookEventArgs> MessageReceived;
        public void Attach()
        {
            if (IsAttached)
            {
                return;
            }

            _hookId = User32.SetWindowsHookEx(WindowsHook.MOUSE_LL, HookCallback, Kernel32.GetModuleHandle("user32"), 0);
            if (_hookId == IntPtr.Zero)
            {
                throw new System.ComponentModel.Win32Exception();
            }

            IsAttached = true;
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                return;
            }

            IsAttached = !User32.UnhookWindowsHookEx(_hookId);
        }

        public void Dispose()
        {
            Detach();
        }
    }
}