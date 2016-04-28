using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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

        //private readonly IAttachable _timerHook;
        private readonly IGeneralConfigService _config;
        private readonly IList<IAttachable> _hooks = new List<IAttachable>();

        public RubberduckHooks(VBE vbe, IGeneralConfigService config)
        {
            _vbe = vbe;
            _mainWindowHandle = (IntPtr)vbe.MainWindow.HWnd;
            _oldWndProc = WindowProc;
            _newWndProc = WindowProc;
            _oldWndPointer = User32.SetWindowLong(_mainWindowHandle, (int)WindowLongFlags.GWL_WNDPROC, _newWndProc);
            _oldWndProc = (User32.WndProc)Marshal.GetDelegateForFunctionPointer(_oldWndPointer, typeof(User32.WndProc));

            _config = config;

        }

        public void HookHotkeys()
        {
            Detach();
            _hooks.Clear();

            var config = _config.LoadConfiguration();
            var settings = config.UserSettings.GeneralSettings.HotkeySettings;

            AddHook(new MouseHook(_mainWindowHandle));
            AddHook(new KeyboardHook(_mainWindowHandle));
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

            try
            {
                foreach (var hook in Hooks)
                {
                    hook.Attach();
                    hook.MessageReceived += hook_MessageReceived;
                }

                IsAttached = true;
            }
            catch (Win32Exception exception)
            {
                Debug.WriteLine(exception);
            }
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                return;
            }

            try
            {
                foreach (var hook in Hooks)
                {
                    hook.MessageReceived -= hook_MessageReceived;
                    hook.Detach();
                }
            }
            catch (Win32Exception exception)
            {
                Debug.WriteLine(exception);
            }
            IsAttached = false;
        }

        private void hook_MessageReceived(object sender, HookEventArgs e)
        {
            if (sender is MouseHook)
            {
                Debug.WriteLine("MouseHook message received");
                OnMessageReceived(sender, e);
                return;
            }

            if (sender is KeyboardHook)
            {
                Debug.WriteLine("KeyboardHook message received");
                OnMessageReceived(sender, e);
                return;
            }

            var hotkey = sender as IHotkey;
            if (hotkey != null)
            {
                Debug.WriteLine("Hotkey message received");
                hotkey.Command.Execute(null);
                return;
            }

            Debug.WriteLine("Unknown message received");
            OnMessageReceived(sender, e);
        }

        public void Dispose()
        {
            Detach();
        }

        private IntPtr WindowProc(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                var suppress = false;
                if (hWnd == _mainWindowHandle)
                {
                    switch ((WM)uMsg)
                    {
                        case WM.HOTKEY:
                            suppress = HandleHotkeyMessage(wParam);
                            break;

                        case WM.ACTIVATEAPP:
                            HandleActivateAppMessage(wParam);
                            break;
                    }
                }

                return suppress 
                    ? IntPtr.Zero 
                    : User32.CallWindowProc(_oldWndProc, hWnd, uMsg, wParam, lParam);
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
            }

            return IntPtr.Zero;
        }

        private bool HandleHotkeyMessage(IntPtr wParam)
        {
            var processed = false;
            try
            {
                if (User32.IsVbeWindowActive(_mainWindowHandle))
                {
                    var hook = Hooks.OfType<Hotkey>().SingleOrDefault(k => k.HotkeyInfo.HookId == wParam);
                    if (hook != null)
                    {
                        hook.OnMessageReceived();
                        processed = true;
                    }
                }
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
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
    }
}
