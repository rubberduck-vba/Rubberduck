using System;
using System.ComponentModel;
using System.Diagnostics;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class MouseHook : IAttachable
    {
        private IntPtr _hookId;

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var button = (WM)wParam;
                if (button == WM.RBUTTONDOWN || button == WM.LBUTTONDOWN)
                {
                    // handle right-click to evaluate commands' CanExecute before the context menu is shown;
                    // handle left-click to do the same before the Rubberduck menu is drawn, too.
                    OnMessageReceived();
                }
            }

            return User32.CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private void OnMessageReceived()
        {
            var handler = MessageReceived;
            if (handler != null)
            {
                handler.Invoke(this, HookEventArgs.Empty);
            }
        }

        public bool IsAttached { get; private set; }
        public event EventHandler<HookEventArgs> MessageReceived;

        public void Attach()
        {
            if (IsAttached)
            {
                return;
            }

            var handle = Kernel32.GetModuleHandle("user32");
            if (handle == IntPtr.Zero)
            {
                throw new Win32Exception();
            } 
            _hookId = User32.SetWindowsHookEx(WindowsHook.MOUSE, HookCallback, handle, 0);
            if (_hookId == IntPtr.Zero)
            {
                throw new Win32Exception();
            }
            
            IsAttached = true;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                return;
            }

            if (!User32.UnhookWindowsHookEx(_hookId))
            {
                throw new Win32Exception();
            }

            IsAttached = false;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }
    }
}