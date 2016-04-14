using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class KeyboardHook : IAttachable
    {
        private readonly VBE _vbe;
        private IntPtr _hookId;

        private readonly User32.HookProc _callback;

        public KeyboardHook(VBE vbe)
        {
            _vbe = vbe;
            _callback = HookCallback;
        }

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

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                var pane = _vbe.ActiveCodePane;
                if ((WM)wParam == WM.KEYUP
                    && pane != null
                    && NavigationKeys.Contains((Keys)Marshal.ReadInt32(lParam)) 
                    && User32.IsVbeWindowActive((IntPtr)_vbe.MainWindow.HWnd))
                {
                    int startLine;
                    int endLine;
                    int startColumn;
                    int endColumn;

                    // not using extension method because a QualifiedSelection would be overkill:
                    pane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                    if (nCode >= 0)
                    {
                        OnMessageReceived();
                    }
                }

                return User32.CallNextHookEx(_hookId, nCode, wParam, lParam);
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
            }

            return IntPtr.Zero;
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

            _hookId = User32.SetWindowsHookEx(WindowsHook.KEYBOARD_LL, _callback, handle, 0);
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