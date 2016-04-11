using System;
using System.ComponentModel;
using System.Diagnostics;
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

        public KeyboardHook(VBE vbe)
        {
            _vbe = vbe;
        }

        private int _lastLineIndex;
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            var pane = _vbe.ActiveCodePane;
            if (pane != null)
            {
                int startLine;
                int endLine;
                int startColumn;
                int endColumn;

                // not using extension method because a QualifiedSelection would be overkill:
                pane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);                
                if (startLine != _lastLineIndex)
                {
                    // if the current line has changed, let the KEYDOWN be written to the IDE, and notify on KEYUP:
                    _lastLineIndex = startLine;
                    if (nCode >= 0 && (WM)wParam == WM.KEYUP)
                    {
                        //var key = (Keys)Marshal.ReadInt32(lParam);
                        OnMessageReceived();
                    }
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

            _hookId = User32.SetWindowsHookEx(WindowsHook.KEYBOARD, HookCallback, Kernel32.GetModuleHandle("user32"), 0);
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

            User32.UnhookWindowsHookEx(_hookId);

            IsAttached = false;
            Debug.WriteLine("{0}: {1}", GetType().Name, IsAttached ? "Attached" : "Detached");
        }
    }
}