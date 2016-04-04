using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Common.WinAPI;
using Rubberduck.Settings;

namespace Rubberduck.Common
{
    public class RubberduckHooks : IRubberduckHooks
    {
        private readonly VBE _vbe;
        private readonly IntPtr _mainWindowHandle;

        private readonly IntPtr _oldWndPointer;
        private readonly User32.WndProc _oldWndProc;
        private User32.WndProc _newWndProc;

        private readonly IAttachable _timerHook;
        private readonly IGeneralConfigService _config;
        private readonly IList<IAttachable> _hooks = new List<IAttachable>();

        public RubberduckHooks(VBE vbe, IAttachable timerHook, IGeneralConfigService config)
        {
            _vbe = vbe;
            _mainWindowHandle = (IntPtr)vbe.MainWindow.HWnd;
            _oldWndProc = WindowProc;
            _newWndProc = WindowProc;
            _oldWndPointer = User32.SetWindowLong(_mainWindowHandle, (int)WindowLongFlags.GWL_WNDPROC, _newWndProc);
            _oldWndProc = (User32.WndProc)Marshal.GetDelegateForFunctionPointer(_oldWndPointer, typeof(User32.WndProc));

            _timerHook = timerHook;
            _config = config;
            _timerHook.MessageReceived += timerHook_MessageReceived;
        }

        public void HookHotkeys()
        {
            Detach();
            _hooks.Clear();

            var config = _config.LoadConfiguration();
            var settings = config.UserSettings.GeneralSettings.HotkeySettings;
            foreach (var hotkey in settings.Where(hotkey => hotkey.IsEnabled))
            {
                AddHook(new Hotkey(_mainWindowHandle, hotkey.ToString(), hotkey.Command));
            }

            Attach();
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
                hook.MessageReceived += hook_MessageReceived;
            }

            IsAttached = true;
        }

        private void hook_MessageReceived(object sender, HookEventArgs e)
        {
            if (sender is ILowLevelKeyboardHook)
            {
                // todo: handle 2-step hotkeys?
                return;
            }

            var hotkey = sender as IHotkey;
            if (hotkey != null)
            {
                hotkey.Command.Execute(null);
            }

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
                hook.MessageReceived -= hook_MessageReceived;
            }

            IsAttached = false;
        }

        private void timerHook_MessageReceived(object sender, EventArgs e)
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
            _timerHook.MessageReceived -= timerHook_MessageReceived;
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
                            processed = HandleHotkeyMessage(wParam);
                            break;

                        case WM.ACTIVATEAPP:
                            HandleActivateAppMessage(wParam);
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

        private bool HandleHotkeyMessage(IntPtr wParam)
        {
            var processed = false;
                            if (GetWindowThread(User32.GetForegroundWindow()) == GetWindowThread(_mainWindowHandle))
                            {
                var hook = Hooks.OfType<Hotkey>().SingleOrDefault(k => k.HotkeyInfo.HookId == wParam);
                                if (hook != null)
                                {
                    hook.OnMessageReceived();
                    processed = true;
                }
            }
            return processed;
        }

        private void HandleActivateAppMessage(IntPtr wParam)
        {
            const int WA_INACTIVE = 0;
            const int WA_ACTIVE = 1;
            const int WA_CLICKACTIVE = 2;

            switch (LoWord(wParam))
            {
                case WA_ACTIVE:
                case WA_CLICKACTIVE:
                    Attach();
                    break;

                case WA_INACTIVE:
                    Detach();
                    break;
            }
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
