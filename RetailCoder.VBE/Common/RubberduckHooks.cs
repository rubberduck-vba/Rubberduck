using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Common.WinAPI;

namespace Rubberduck.Common
{
    public class RubberduckHooks : IRubberduckHooks
    {
        private readonly IntPtr _mainWindowHandle;

        private readonly IntPtr _oldWndPointer;
        private readonly User32.WndProc _oldWndProc;
        private User32.WndProc _newWndProc;

        private readonly ITimerHook _timerHook;
        private readonly IList<IAttachable> _hooks = new List<IAttachable>();


        private const int WA_INACTIVE = 0;
        private const int WA_ACTIVE = 1;

        public RubberduckHooks(IntPtr mainWindowHandle, ITimerHook timerHook)
        {
            _mainWindowHandle = mainWindowHandle;
            _oldWndProc = WindowProc;
            _newWndProc = WindowProc;
            _oldWndPointer = User32.SetWindowLong(_mainWindowHandle, (int)WindowLongFlags.GWL_WNDPROC, _newWndProc);
            _oldWndProc = (User32.WndProc)Marshal.GetDelegateForFunctionPointer(_oldWndPointer, typeof(User32.WndProc));

            _timerHook = timerHook;
            _timerHook.Tick += timerHook_Tick;
        }

        public IEnumerable<IAttachable> Hooks { get { return _hooks; } }

        public void AddHook(IAttachable hook)
        {
            _hooks.Add(hook);
        }

        public event EventHandler<HookEventArgs> MessageReceived;

        private void OnMessageReceived(object sender, HookEventArgs args)
        {
            var handler = MessageReceived;
            if (handler != null)
            {
                handler.Invoke(sender, args);
            }
        }

        public bool IsAttached { get; private set; }

        public void Attach()
        {
            if (IsAttached)
            {
                return;
            }

            foreach (var hook in Hooks)
            {
                hook.Attach();
                var h = hook as IHook;
                if (h != null)
                {
                    h.MessageReceived += hook_MessageReceived;
                }
            }

            IsAttached = true;
        }

        private void hook_MessageReceived(object sender, HookEventArgs e)
        {
            OnMessageReceived(sender, e);
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                return;
            }

            foreach (var hook in Hooks)
            {
                hook.Detach();
                var h = hook as IHook;
                if (h != null)
                {
                    h.MessageReceived -= hook_MessageReceived;
                }
            }

            IsAttached = false;
        }

        private void timerHook_Tick(object sender, EventArgs e)
        {
            if (!IsAttached && User32.GetForegroundWindow() == _mainWindowHandle)
            {
                Attach();
            }
            else
            {
                Detach();
            }
        }

        public void Dispose()
        {
            _timerHook.Tick -= timerHook_Tick;
            _timerHook.Detach();

            Detach();
        }

        private IntPtr WindowProc(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                var processed = false;
                if (hWnd == _mainWindowHandle)
                {
                    switch ((WM)uMsg)
                    {
                        case WM.HOTKEY:
                            if (GetWindowThread(User32.GetForegroundWindow()) == GetWindowThread(_mainWindowHandle))
                            {
                                var hook = Hooks.OfType<IHotkey>().SingleOrDefault(k => k.HotkeyInfo.HookId == (IntPtr)wParam);
                                if (hook != null)
                                {
                                    var args = new HookEventArgs(hook.HotkeyInfo.Keys);
                                    OnMessageReceived(hook, args);
                                    processed = true;
                                }
                            }
                            break;

                        case WM.ACTIVATEAPP:
                            switch (LoWord(wParam))
                            {
                                case WA_ACTIVE:
                                    Attach();
                                    break;

                                case WA_INACTIVE:
                                    Detach();
                                    break;
                            }

                            break;
                    }
                }

                if (!processed)
                {
                    return User32.CallWindowProc(_oldWndProc, hWnd, uMsg, wParam, lParam);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }

            return User32.CallWindowProc(_oldWndProc, hWnd, uMsg, wParam, lParam);
        }

        private static int LoWord(IntPtr dw)
        {
            return unchecked((short)(uint)dw);
        }

        private IntPtr GetWindowThread(IntPtr hWnd)
        {
            uint hThread;
            User32.GetWindowThreadProcessId(hWnd, out hThread);

            return (IntPtr)hThread;
        }
    }
}