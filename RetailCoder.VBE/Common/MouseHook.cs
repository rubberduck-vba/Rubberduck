using System;
using System.ComponentModel;
using System.Diagnostics;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class MouseHook : IAttachable
    {
        private readonly VBE _vbe;
        private IntPtr _hookId;
        private readonly User32.HookProc _callback;

        public MouseHook(VBE vbe)
        {
            _vbe = vbe;
            _callback = HookCallback;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                var pane = _vbe.ActiveCodePane;
                if (User32.IsVbeWindowActive((IntPtr)_vbe.MainWindow.HWnd) && nCode >= 0 && pane != null)
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
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
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
            _hookId = User32.SetWindowsHookEx(WindowsHook.MOUSE_LL, _callback, handle, 0);
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

            IsAttached = false;
            if (!User32.UnhookWindowsHookEx(_hookId))
            {
                _hookId = IntPtr.Zero;
                throw new Win32Exception();
            }

            _hookId = IntPtr.Zero;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }
    }
}