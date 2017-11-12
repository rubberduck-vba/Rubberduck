using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Settings;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.Common
{
    public class RubberduckHooks : SubclassingWindow, IRubberduckHooks
    {
        private readonly IGeneralConfigService _config;
        private readonly IList<IAttachable> _hooks = new List<IAttachable>();
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public RubberduckHooks(IVBE vbe, IGeneralConfigService config)
            : base((IntPtr)vbe.MainWindow.HWnd, (IntPtr)vbe.MainWindow.HWnd)
        {
            //_commands = commands;
            _config = config;
        }

        public void HookHotkeys()
        {
            Detach();
            _hooks.Clear();

            var config = _config.LoadConfiguration();
            var settings = config.UserSettings.HotkeySettings;

            foreach (var hotkey in settings.Settings.Where(hotkey => hotkey.IsEnabled))
            {
                var command = hotkey.Command;

                // TODO: Is this check needed?
                if (command != null)
                {
                    AddHook(new Hotkey(Hwnd, hotkey.ToString(), command));
                }
            }

            Attach();
        }

        public IEnumerable<IAttachable> Hooks => _hooks;

        public void AddHook(IAttachable hook)
        {
            _hooks.Add(hook);
        }

        public event EventHandler<HookEventArgs> MessageReceived;

        private void OnMessageReceived(object sender, HookEventArgs args)
        {
            var handler = MessageReceived;
            handler?.Invoke(sender, args);
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
                Logger.Error(exception);
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
                Logger.Error(exception);
            }
            IsAttached = false;
        }

        private void hook_MessageReceived(object sender, HookEventArgs e)
        {
            if (sender is IHotkey hotkey && hotkey.Command.CanExecute(null))
            {
                hotkey.Command.Execute(null);
                return;
            }
            
            OnMessageReceived(sender, e);
        }

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            var suppress = false;
            switch ((WM)msg)
            {
                case WM.HOTKEY:
                    suppress = hWnd == Hwnd && HandleHotkeyMessage(wParam);
                    break;
                case WM.SETFOCUS:
                    Attach();
                    break;
                case WM.RUBBERDUCK_CHILD_FOCUS:
                    if (lParam == IntPtr.Zero)
                    {
                        Detach();
                    }
                    else
                    {
                        Attach();
                    }
                    suppress = true;
                    break;
                case WM.NCACTIVATE:
                    if (wParam == IntPtr.Zero)
                    {
                        Detach();
                    }
                    break;
                case WM.CLOSE:
                case WM.DESTROY:
                case WM.RUBBERDUCK_SINKING:
                    Detach();
                    break;
            }
            return suppress ? 0 : base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }

        private bool HandleHotkeyMessage(IntPtr wParam)
        {
            var processed = false;
            try
            {
                var hook = Hooks.OfType<Hotkey>().SingleOrDefault(k => k.HotkeyInfo.HookId == wParam);
                if (hook != null)
                {
                    hook.OnMessageReceived();
                    processed = true;
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception);
            }
            return processed;
        }
    }
}
