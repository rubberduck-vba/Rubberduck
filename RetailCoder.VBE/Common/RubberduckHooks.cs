using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
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
        private readonly IEnumerable<ICommand> _commands;
        private readonly IList<IAttachable> _hooks = new List<IAttachable>();
        private readonly IDictionary<RubberduckHotkey, ICommand> _mappings;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RubberduckHooks(VBE vbe, IGeneralConfigService config, IEnumerable<ICommand> commands)
        {
            _vbe = vbe;
            _mainWindowHandle = (IntPtr)vbe.MainWindow.HWnd;
            _oldWndProc = WindowProc;
            _newWndProc = WindowProc;
            _oldWndPointer = User32.SetWindowLong(_mainWindowHandle, (int)WindowLongFlags.GWL_WNDPROC, _newWndProc);
            _oldWndProc = (User32.WndProc)Marshal.GetDelegateForFunctionPointer(_oldWndPointer, typeof(User32.WndProc));

            _commands = commands;
            _config = config;
            _mappings = GetCommandMappings();
        }

        private ICommand Command<TCommand>() where TCommand : ICommand
        {
            return _commands.OfType<TCommand>().SingleOrDefault();
        }

        private IDictionary<RubberduckHotkey, ICommand> GetCommandMappings()
        {
            return new Dictionary<RubberduckHotkey, ICommand>
            {
                { RubberduckHotkey.ParseAll, Command<ReparseCommand>() },
                { RubberduckHotkey.CodeExplorer, Command<CodeExplorerCommand>() },
                { RubberduckHotkey.IndentModule, Command<IndentCurrentModuleCommand>() },
                { RubberduckHotkey.IndentProcedure, Command<IndentCurrentProcedureCommand>() },
                { RubberduckHotkey.FindSymbol, Command<FindSymbolCommand>() },
                { RubberduckHotkey.RefactorMoveCloserToUsage, Command<RefactorMoveCloserToUsageCommand>() },
                { RubberduckHotkey.InspectionResults, Command<InspectionResultsCommand>() },
                { RubberduckHotkey.RefactorExtractMethod, Command<RefactorExtractMethodCommand>() },
                { RubberduckHotkey.RefactorRename, Command<CodePaneRefactorRenameCommand>() },
                { RubberduckHotkey.TestExplorer, Command<TestExplorerCommand>() }
            };
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
                    AddHook(new Hotkey(_mainWindowHandle, hotkey.ToString(), _mappings[assigned]));
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

        private void Keyboard_RawKeyboardInputReceived(object sender, RawKeyEventArgs e)
        {
            if (e.Message == WM.KEYUP)
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
                _logger.Error(exception, "Attaching hooks failed.");
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
                _logger.Error(exception, "Detaching hooks failed.");
            }
            IsAttached = false;
        }

        private void hook_MessageReceived(object sender, HookEventArgs e)
        {
            var hotkey = sender as IHotkey;
            if (hotkey != null)
            {
                _logger.Debug("Hotkey message received");                
                hotkey.Command.Execute(null);
                return;
            }

            _logger.Debug("Unknown message received");
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
                _logger.Error(exception, "Encountered error in WindowProc.");
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
                _logger.Error(exception, "Encountered error in HandleHotkeyMessage.");
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
