using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class LowLevelKeyboardHook : ILowLevelKeyboardHook, IDisposable
    {
        private readonly VBE _vbe;
        private readonly IntPtr _hWndVbe;

        private readonly User32.HookProc _proc;
        private static IntPtr _hookId = IntPtr.Zero;

        public LowLevelKeyboardHook(VBE vbe)
        {
            _vbe = vbe;
            _hWndVbe = (IntPtr)_vbe.MainWindow.HWnd;

            _proc = HookCallback;
        }

        public event EventHandler<HookEventArgs> MessageReceived;
        public void OnMessageReceived()
        {
            var handler = MessageReceived;
            if (handler != null)
            {
                handler.Invoke(this, HookEventArgs.Empty);
            }
        }

        public bool IsAttached { get; private set; }

        public bool EatNextKey { get; set; }

        private static readonly Keys[] IgnoredKeys = 
        {
            Keys.Down,
            Keys.Up,
            Keys.Left,
            Keys.Right,
            Keys.PageDown,
            Keys.PageUp,
            Keys.CapsLock,
            Keys.Escape,
            Keys.Home,
            Keys.End,
            Keys.Shift,
            Keys.ShiftKey,
            Keys.LShiftKey,
            Keys.RShiftKey,
            Keys.Control,
            Keys.ControlKey,
            Keys.LControlKey,
            Keys.RControlKey,
        };

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (_vbe.ActiveWindow == null || _vbe.ActiveWindow.Type != vbext_WindowType.vbext_wt_CodeWindow)
            {
                // don't do anything if not in a code window
                return User32.CallNextHookEx(_hookId, nCode, wParam, lParam);
            }

            var vkCode = Marshal.ReadInt32(lParam);
            var key = (Keys)vkCode;

            var windowHandle = User32.GetForegroundWindow();
            var codePane = _vbe.ActiveCodePane;

            Task.Run(() =>
            {
                if (windowHandle != _hWndVbe
                    || wParam != (IntPtr) WM.KEYUP
                    || nCode < 0 
                    || codePane == null
                    || IgnoredKeys.Contains(key))
                {
                    return;
                }

                OnMessageReceived();
            });

            return EatNextKey ? (IntPtr)1 : User32.CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        public void Attach()
        {
            if (IsAttached)
            {
                return;
            }

            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                _hookId = User32.SetWindowsHookEx(WindowsHook.KEYBOARD_LL, _proc, Kernel32.GetModuleHandle(curModule.ModuleName), 0);
                IsAttached = true;
            }
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                return;
            }

            User32.UnhookWindowsHookEx(_hookId);
            IsAttached = false;
        }

        public event EventHandler<KeyHookEventArgs> KeyPressed;

        private void OnKeyPressed(KeyHookEventArgs e)
        {
            var handler = KeyPressed;
            if (handler != null)
            {
                handler.Invoke(this, e);
            }
        }

        public void Dispose()
        {
            Detach();
        }
    }
}
