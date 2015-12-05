using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Castle.DynamicProxy.Contributors;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public interface IKeyHook
    {
        void Attach();
        void Detach();
        event EventHandler<KeyHookEventArgs> KeyPressed;
    }

    public class KeyHook : IKeyHook, IDisposable
    {
        private readonly VBE _vbe;

        private readonly HashSet<HookInfo> _hookedKeys = new HashSet<HookInfo>();

        private const int GWL_WNDPROC = -4;
        private const int WA_INACTIVE = 0;
        private const int WA_ACTIVE = 1;

        private User32.WndProc _oldWndProc;
        private IntPtr _hWndForm;
        private IntPtr _hWndVbe;

        private bool _isRegistered = false;

        private readonly User32.HookProc _proc;
        private static IntPtr HookId = IntPtr.Zero;


        private static IntPtr SetHook(User32.HookProc proc)
        {
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return User32.SetWindowsHookEx(WindowsHook.KEYBOARD_LL, proc, Kernel32.GetModuleHandle(curModule.ModuleName), 0);
            }
        }

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
            if (_vbe.ActiveWindow.Type != vbext_WindowType.vbext_wt_CodeWindow)
            {
                // don't do anything if not in a code window
                return User32.CallNextHookEx(HookId, nCode, wParam, lParam);
            }

            var vkCode = Marshal.ReadInt32(lParam);
            var key = (Keys)vkCode;

            _hWndVbe = (IntPtr)_vbe.MainWindow.HWnd;
            var windowHandle = User32.GetForegroundWindow();
            var codePane = _vbe.ActiveCodePane;

            Task.Run(() =>
            {
                if (windowHandle != _hWndForm
                    || wParam != (IntPtr) WM.KEYUP 
                    || nCode < 0 
                    || codePane == null
                    || IgnoredKeys.Contains(key))
                {
                    return;
                }

                var component = codePane.CodeModule.Parent;
                var args = new KeyHookEventArgs(key, component);
                OnKeyPressed(args);
            });

            return User32.CallNextHookEx(HookId, nCode, wParam, lParam);
        }

        public KeyHook(VBE vbe)
        {
            _vbe = vbe;
            _proc = HookCallback;
        }

        public void Attach()
        {
            HookId = SetHook(_proc);
        }

        public void Detach()
        {
            User32.UnhookWindowsHookEx(HookId);
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

        private static readonly IDictionary<string, Keys> _keys = new Dictionary<string, Keys>
        {
            { "{BACKSPACE}", Keys.Back },
            { "{BS}", Keys.Back },
            { "{BKSP}", Keys.Back },
            { "{CAPSLOCK}", Keys.CapsLock },
            { "{DELETE}", Keys.Delete },
            { "{DEL}", Keys.Delete },
            { "{DOWN}", Keys.Down },
            { "{END}", Keys.End },
            { "{ENTER}", Keys.Enter },
            { "{RETURN}", Keys.Enter },
            { "{ESC}", Keys.Escape },
            { "{HELP}", Keys.Help },
            { "{HOME}", Keys.Home },
            { "{INSERT}", Keys.Insert },
            { "{INS}", Keys.Insert },
            { "{LEFT}", Keys.Left },
            { "{NUMLOCK}", Keys.NumLock },
            { "{PGDN}", Keys.PageDown },
            { "{PGUP}", Keys.PageUp },
            { "{PRTSC}", Keys.PrintScreen },
            { "{RIGHT}", Keys.Right },
            { "{TAB}", Keys.Tab },
            { "{UP}", Keys.Up },
            { "{F1}", Keys.F1 },
            { "{F2}", Keys.F2 },
            { "{F3}", Keys.F3 },
            { "{F4}", Keys.F4 },
            { "{F5}", Keys.F5 },
            { "{F6}", Keys.F6 },
            { "{F7}", Keys.F7 },
            { "{F8}", Keys.F8 },
            { "{F9}", Keys.F9 },
            { "{F10}", Keys.F10 },
            { "{F11}", Keys.F11 },
            { "{F12}", Keys.F12 },
            { "{F13}", Keys.F13 },
            { "{F14}", Keys.F14 },
            { "{F15}", Keys.F15 },
            { "{F16}", Keys.F16 },
        };

        private IntPtr _timerId;

        private Keys GetKey(string keyCode)
        {
            var result = Keys.None;
            switch (keyCode.Substring(0,1))
            {
                case "{":
                    _keys.TryGetValue(keyCode, out result);
                    break;
                case "~":
                    result = Keys.Return;
                    break;
                default:
                    if (!string.IsNullOrEmpty(keyCode))
                    {
                        int asciiCode;
                        if (int.TryParse(keyCode.Substring(0, 1), out asciiCode))
                        {
                            result = (Keys)asciiCode;
                        }
                    }
                    break;
            }

            return result;
        }

        private void HookKey(int keyCode, int shift, Action action)
        {
            UnHookKey(keyCode, shift);

            if (!_hookedKeys.Any())
            {
                // HookWindow();
            }

            var hookId = (IntPtr)Kernel32.GlobalAddAtom(Guid.NewGuid().ToString());
            User32.RegisterHotKey((IntPtr)_vbe.MainWindow.HWnd, hookId, (uint)shift, (uint)keyCode);

            _hookedKeys.Add(new HookInfo(hookId, (uint)keyCode, shift, action));
        }

        private void UnHookKey(int keyCode, int shift)
        {
            var hooks = _hookedKeys.Where(hook => hook.KeyCode == keyCode && hook.Shift == shift).ToList();
            foreach (var hook in hooks)
            {
                User32.UnregisterHotKey((IntPtr)_vbe.MainWindow.HWnd, hook.HookId);
                Kernel32.GlobalDeleteAtom((ushort)hook.HookId);
                _hookedKeys.Remove(hook);

                // if (!_hookedKeys.Any()) { UnHookWindow(); }
            }
        }

        public void UnhookAll()
        {
            foreach (var hook in _hookedKeys)
            {
                User32.UnregisterHotKey((IntPtr)_vbe.MainWindow.HWnd, hook.HookId);
                Kernel32.GlobalDeleteAtom((ushort)hook.HookId);
            }

            //UnHookWindow();
        }

        private void HookWindow()
        {
            var hwnd = (IntPtr)_vbe.MainWindow.HWnd; // use OnKeyWindow?
            _oldWndProc = (User32.WndProc)Marshal.GetDelegateForFunctionPointer(
                    User32.SetWindowLongPtr(hwnd, WindowLongFlags.GWL_WNDPROC,
                    Marshal.GetFunctionPointerForDelegate((User32.WndProc) WindowProc)), typeof (User32.WndProc));

        }

        private void UnHookWindow()
        {
            var hwnd = (IntPtr)_vbe.MainWindow.HWnd; // use OnKeyWindow?
            User32.SetWindowLongPtr(hwnd, WindowLongFlags.GWL_WNDPROC, Marshal.GetFunctionPointerForDelegate(_oldWndProc));
            hwnd = IntPtr.Zero;
            _isRegistered = false;

            User32.KillTimer(hwnd, _timerId);
            Kernel32.GlobalDeleteAtom((ushort)_timerId);
            _timerId = IntPtr.Zero;

            // dispose OnKeyWindow instance
        }

        private IntPtr WindowProc(IntPtr hWnd, uint u, IntPtr wParam, IntPtr lParam)
        {
            var processed = false;
            if (hWnd == _hWndForm)
            {
                switch ((WM)u)
                {
                    case WM.HOTKEY:
                        // if (GetWindowThread(User32.GetForegroundWindow()) == GetWindowThread(_hWndVBE))
                        {
                            var key = _hookedKeys.FirstOrDefault(k => (Keys) k.KeyCode == (Keys) wParam);
                            key.Action.Invoke();
                            processed = true;
                        }
                        break;

                    case WM.ACTIVATEAPP:
                        switch (LoWord((int)wParam))
                        {
                            case WA_ACTIVE:
                                foreach (var key in _hookedKeys)
                                {
                                    User32.RegisterHotKey(_hWndForm, key.HookId, User32.MOD_SHIFT | User32.MOD_CONTROL, key.KeyCode);
                                }
                                _isRegistered = true;
                                break;

                            case WA_INACTIVE:
                                foreach (var key in _hookedKeys)
                                {
                                    User32.UnregisterHotKey(_hWndForm, key.HookId);
                                }
                                _isRegistered = false;
                                break;
                        }

                        break;
                }
            }

            if (!processed)
            {
                return User32.CallWindowProc(_oldWndProc, hWnd, u, wParam, lParam);
            }
            return IntPtr.Zero;
        }

        private int LoWord(int dw)
        {
            return (dw & 0x8000) != 0 
                ? 0x8000 | (dw & 0x7FFF) 
                : dw & 0xFFFF;
        }

        private void TimerCallback(IntPtr hWnd, WindowLongFlags msg, IntPtr timerId, IntPtr time)
        {
            
        }


    }
}
