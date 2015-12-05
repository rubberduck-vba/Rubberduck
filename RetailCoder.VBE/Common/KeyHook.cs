using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public interface IKeyHook
    {
        /// <summary>
        /// Attaches a system keyhook that listens to all keypresses.
        /// </summary>
        void Attach();
        /// <summary>
        /// Detaches system keyhook.
        /// </summary>
        void Detach();
        /// <summary>
        /// Raised when system keyhook captures a keypress in the VBE.
        /// </summary>
        event EventHandler<KeyHookEventArgs> KeyPressed;
        /// <summary>
        /// Registers specified delegate for specified key combination.
        /// </summary>
        /// <param name="key">The key combination string, including modifiers ('+': Shift, '%': Alt, '^': Control).</param>
        /// <param name="action">Any <c>void</c>, parameterless method that handles the hotkey.</param>
        void OnHotKey(string key, Action action = null);
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
        private VbeOnKey _hookForm = new VbeOnKey();

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

        public void OnHotKey(string key, Action action = null)
        {
            var hotKey = key;
            var lShift = GetModifierValue(ref hotKey);
            var lKey = GetKey(hotKey);

            if (lKey == Keys.None)
            {
                throw new InvalidOperationException("Invalid key.");
            }

            if (action == null)
            {
                UnHookKey((uint)lKey, lShift);
            }
            else
            {
                HookKey((uint)lKey, lShift, action);
            }

            _hWndVbe = (IntPtr)_vbe.MainWindow.HWnd;
        }

        /// <summary>
        /// Gets the <see cref="KeyModifier"/> values out of a key combination.
        /// </summary>
        /// <param name="key">The hotkey string, returned without the modifiers.</param>
        private static uint GetModifierValue(ref string key)
        {
            uint lShift = 0;
            for (var i = 0; i < 3; i++)
            {
                var firstChar = key.Substring(0, 1);
                if (firstChar == "+")
                {
                    lShift |= (uint)KeyModifier.SHIFT;
                }
                else if (firstChar == "%")
                {
                    lShift |= (uint)KeyModifier.ALT;
                }
                else if (firstChar == "^")
                {
                    lShift |= (uint)KeyModifier.CONTROL;
                }
                else
                {
                    break;
                }

                key = key.Substring(1);
            }
            return lShift;
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

        private void HookKey(uint keyCode, uint shift, Action action)
        {
            UnHookKey(keyCode, shift);

            if (!_hookedKeys.Any())
            {
                HookWindow();
            }

            var hookId = (IntPtr)Kernel32.GlobalAddAtom(Guid.NewGuid().ToString());
            User32.RegisterHotKey(_hWndForm, hookId, shift, keyCode);

            _hookedKeys.Add(new HookInfo(hookId, keyCode, shift, action));
            _isRegistered = true;
        }

        private void UnHookKey(uint keyCode, uint shift)
        {
            var hooks = _hookedKeys.Where(hook => hook.KeyCode == keyCode && hook.Shift == shift).ToList();
            foreach (var hook in hooks)
            {
                User32.UnregisterHotKey(_hWndForm, hook.HookId);
                Kernel32.GlobalDeleteAtom((ushort)hook.HookId);
                _hookedKeys.Remove(hook);

                if (!_hookedKeys.Any())
                {
                    UnHookWindow();
                    break;
                }
            }
        }

        /// <summary>
        /// Called when hook form goes out of scope, to remove all hooks.
        /// </summary>
        public void UnHookAll()
        {
            foreach (var hook in _hookedKeys)
            {
                User32.UnregisterHotKey(_hWndForm, hook.HookId);
                Kernel32.GlobalDeleteAtom((ushort)hook.HookId);
            }

            UnHookWindow();
            Detach();
        }

        private void HookWindow()
        {
            try
            {
                _hookForm = new VbeOnKey();
                _hookForm.Closed += _hookForm_Closed;
                _hWndForm = _hookForm.Handle;
                _oldWndProc = (User32.WndProc)Marshal.GetDelegateForFunctionPointer(
                    User32.SetWindowLongPtr(_hWndForm, WindowLongFlags.GWL_WNDPROC,
                        Marshal.GetFunctionPointerForDelegate((User32.WndProc) WindowProc)), typeof (User32.WndProc));

                var timerId = (IntPtr)Kernel32.GlobalAddAtom(Guid.NewGuid().ToString());
                User32.SetTimer(_hWndForm, timerId, 500, TimerCallback);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        private void _hookForm_Closed(object sender, EventArgs e)
        {
            UnHookAll();
        }

        private void UnHookWindow()
        {
            try
            {
                User32.SetWindowLongPtr(_hWndForm, WindowLongFlags.GWL_WNDPROC, Marshal.GetFunctionPointerForDelegate(_oldWndProc));
                _isRegistered = false;

                User32.KillTimer(_hWndForm, _timerId);
                Kernel32.GlobalDeleteAtom((ushort)_timerId);

                _timerId = IntPtr.Zero;
                _hWndForm = IntPtr.Zero;

                _hookForm.Close();
                _hookForm.Closed -= _hookForm_Closed;
                _hookForm.Dispose();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        private IntPtr WindowProc(IntPtr hWnd, uint u, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                var processed = false;
                if (hWnd == _hWndForm)
                {
                    switch ((WM)u)
                    {
                        case WM.HOTKEY:
                            if (GetWindowThread(User32.GetForegroundWindow()) == GetWindowThread(_hWndVbe))
                            {
                                var key = _hookedKeys.FirstOrDefault(k => (Keys)k.KeyCode == (Keys)wParam);
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
                                        User32.RegisterHotKey(_hWndForm, key.HookId, (uint)KeyModifier.CONTROL, key.KeyCode);
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
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }

            return IntPtr.Zero;
        }

        private IntPtr GetWindowThread(IntPtr hWnd)
        {
            uint hThread;
            User32.GetWindowThreadProcessId(hWnd, out hThread);

            return (IntPtr)hThread;
        }

        /// <summary>
        /// Gets the integer portion of a word
        /// </summary>
        private static int LoWord(int dw)
        {
            return (dw & 0x8000) != 0 
                ? 0x8000 | (dw & 0x7FFF) 
                : dw & 0xFFFF;
        }

        private void TimerCallback(IntPtr hWnd, WindowLongFlags msg, IntPtr timerId, uint time)
        {
            // check if the VBE is still in the foreground
            if (User32.GetForegroundWindow() == _hWndVbe && !_isRegistered)
            {
                // app got focus, re-register hotkeys and re-attach key hook
                foreach (var key in _hookedKeys)
                {
                    User32.RegisterHotKey(_hWndForm, key.HookId, key.Shift, key.KeyCode);
                }
                _isRegistered = true;
                Attach();
            }
            else
            {
                // app lost focus, unregister hotkeys and detach key hook
                foreach (var key in _hookedKeys)
                {
                    User32.UnregisterHotKey(_hWndForm, key.HookId);
                }
                _isRegistered = false;
                Detach();
            }
        }
    }
}
