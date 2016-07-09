using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Common.WinAPI;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using NLog;

namespace Rubberduck.Common
{
    public class RubberduckHooks : IRubberduckHooks
    {
        private readonly VBE _vbe;
        private readonly IntPtr _mainWindowHandle;
        private readonly IntPtr _oldWndPointer;
        private readonly User32.WndProc _oldWndProc;
        private User32.WndProc _newWndProc;
        private RawInput _rawinput;
        private IRawDevice _kb;
        private IRawDevice _mouse;
        private readonly IGeneralConfigService _config;
        private readonly IEnumerable<CommandBase> _commands;
        private readonly IList<IAttachable> _hooks = new List<IAttachable>();
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public RubberduckHooks(VBE vbe, IGeneralConfigService config, IEnumerable<CommandBase> commands)
        {
            _vbe = vbe;
            _mainWindowHandle = (IntPtr)vbe.MainWindow.HWnd;
            _oldWndProc = WindowProc;
            _newWndProc = WindowProc;
            _oldWndPointer = User32.SetWindowLong(_mainWindowHandle, (int)WindowLongFlags.GWL_WNDPROC, _newWndProc);
            _oldWndProc = (User32.WndProc)Marshal.GetDelegateForFunctionPointer(_oldWndPointer, typeof(User32.WndProc));

            _commands = commands;
            _config = config;
        }

        private CommandBase Command(RubberduckHotkey hotkey)
        {
            return _commands.SingleOrDefault(s => s.Hotkey == hotkey);
        }

        public void HookHotkeys()
        {
            Detach();
            _hooks.Clear();

            var config = _config.LoadConfiguration();
            var settings = config.UserSettings.HotkeySettings;

            _rawinput = new RawInput(_mainWindowHandle);

            var kb = (RawKeyboard)_rawinput.CreateKeyboard();
            _rawinput.AddDevice(kb);
            kb.RawKeyInputReceived += Keyboard_RawKeyboardInputReceived;
            _kb = kb;

            var mouse = (RawMouse)_rawinput.CreateMouse();
            _rawinput.AddDevice(mouse);
            mouse.RawMouseInputReceived += Mouse_RawMouseInputReceived;
            _mouse = mouse;

            foreach (var hotkey in settings.Settings.Where(hotkey => hotkey.IsEnabled))
            {
                RubberduckHotkey assigned;
                if (Enum.TryParse(hotkey.Name, out assigned))
                {
                    var command = Command(assigned);
                    Debug.Assert(command != null);

                    AddHook(new Hotkey(_mainWindowHandle, hotkey.ToString(), command));
                }
            }
            Attach();
        }

        private void Mouse_RawMouseInputReceived(object sender, RawMouseEventArgs e)
        {
            if (e.UlButtons.HasFlag(UsButtonFlags.RI_MOUSE_LEFT_BUTTON_UP) || e.UlButtons.HasFlag(UsButtonFlags.RI_MOUSE_RIGHT_BUTTON_UP))
            {
                OnMessageReceived(this, HookEventArgs.Empty);
            }
        }

        // keys that change the current selection.
        private static readonly HashSet<Keys> NavKeys = new HashSet<Keys>
        {
            Keys.Up, Keys.Down, Keys.Left, Keys.Right, Keys.PageDown, Keys.PageUp, Keys.Enter
        };

        private void Keyboard_RawKeyboardInputReceived(object sender, RawKeyEventArgs e)
        {
            // note: handling *all* keys causes annoying RTrim of current line, making editing code a PITA.
            if (e.Message == WM.KEYUP && NavKeys.Contains((Keys)e.VKey))
            {
                OnMessageReceived(this, HookEventArgs.Empty);
            }
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
            var hotkey = sender as IHotkey;
            if (hotkey != null)
            {
                hotkey.Command.Execute(null);
                return;
            }
            
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
                switch ((WM) uMsg)
                {
                    case WM.HOTKEY:
                        suppress = hWnd == _mainWindowHandle && HandleHotkeyMessage(wParam);
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
                        return IntPtr.Zero;
                    case WM.NCACTIVATE:                   
                        if (wParam == IntPtr.Zero)
                        {
                            Detach();
                        }
                        break;
                    case WM.CLOSE:
                        Detach();
                        break;
                }

                return suppress 
                    ? IntPtr.Zero 
                    : User32.CallWindowProc(_oldWndProc, hWnd, uMsg, wParam, lParam);
            }
            catch (Exception exception)
            {
                Logger.Error(exception);
            }

            return IntPtr.Zero;
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
