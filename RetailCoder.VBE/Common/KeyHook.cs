using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Common
{
    public interface IKeyHook
    {
        void Attach();
        void Detach();
        event EventHandler<KeyHookEventArgs> KeyPressed;
    }

    internal struct HookInfo
    {
        private readonly IntPtr _hookId;
        private readonly int _keyCode;
        private readonly int _shift;
        private readonly Action _action;

        public HookInfo(IntPtr hookId, int keyCode, int shift, Action action)
        {
            _hookId = hookId;
            _keyCode = keyCode;
            _shift = shift;
            _action = action;
        }

        public IntPtr HookId { get { return _hookId; } }
        public int KeyCode { get { return _keyCode; } }
        public int Shift { get { return _shift; } }
        public Action Action { get { return _action; } }
    }

    public class KeyHook : IKeyHook, IDisposable
    {
        private readonly VBE _vbe;
        // reference: http://blogs.msdn.com/b/toub/archive/2006/05/03/589423.aspx

        private readonly HashSet<HookInfo> _hookedKeys = new HashSet<HookInfo>();

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;

        private readonly LowLevelKeyboardProc _proc;
        private static IntPtr HookId = IntPtr.Zero;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr RegisterHotKey(int hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr UnRegisterHotKey(int hWnd, int id);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GlobalAddAtom(string lpString);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GlobalDeleteAtom(int nAtom);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr GetWindowThreadProcessId(IntPtr handle, out int processID);

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
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
                return CallNextHookEx(HookId, nCode, wParam, lParam);
            }

            var vkCode = Marshal.ReadInt32(lParam);
            var key = (Keys)vkCode;

            var windowHandle = GetForegroundWindow();
            var vbeWindow = _vbe.MainWindow.HWnd;
            var codePane = _vbe.ActiveCodePane;

            Task.Run(() =>
            {
                if (windowHandle != (IntPtr) vbeWindow 
                    || wParam != (IntPtr) WM_KEYUP 
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

            return CallNextHookEx(HookId, nCode, wParam, lParam);
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
            UnhookWindowsHookEx(HookId);
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

            var hookId = GlobalAddAtom(Guid.NewGuid().ToString());
            RegisterHotKey(_vbe.MainWindow.HWnd, (int)hookId, shift, keyCode);

            _hookedKeys.Add(new HookInfo(hookId, keyCode, shift, action));
        }

        private void UnHookKey(int keyCode, int shift)
        {
            var hooks = _hookedKeys.Where(hook => hook.KeyCode == keyCode && hook.Shift == shift).ToList();
            foreach (var hook in hooks)
            {
                UnRegisterHotKey(_vbe.MainWindow.HWnd, (int)hook.HookId);
                GlobalDeleteAtom((int)hook.HookId);
                _hookedKeys.Remove(hook);

                // if (!_hookedKeys.Any()) { UnHookWindow(); }
            }
        }

        public void UnhookAll()
        {
            foreach (var hook in _hookedKeys)
            {
                UnRegisterHotKey(_vbe.MainWindow.HWnd, (int)hook.HookId);
                GlobalDeleteAtom((int)hook.HookId);
            }

            //UnHookWindow();
        }
    }
}
