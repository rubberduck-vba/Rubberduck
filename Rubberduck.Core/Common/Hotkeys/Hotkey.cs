using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Common.WinAPI;
using NLog;
using Rubberduck.Resources;
using Rubberduck.UI.Command;

namespace Rubberduck.Common.Hotkeys
{
    public class Hotkey : IHotkey
    {
        private readonly IntPtr _hWndVbe;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public Hotkey(IntPtr hWndVbe, string key, CommandBase command, Keys secondKey = Keys.None)
        {
            _hWndVbe = hWndVbe;

            IsTwoStepHotkey = secondKey != Keys.None;
            Key = key;
            Command = command;
            Combo = GetCombo(key);
            SecondKey = secondKey;
        }

        public CommandBase Command { get; }

        public string Key { get; }

        public HotkeyInfo HotkeyInfo { get; private set; }
        public Keys Combo { get; }
        public Keys SecondKey { get; }
        public bool IsTwoStepHotkey { get; }
        public bool IsAttached => HotkeyInfo.HookId != 0;

        public event EventHandler<HookEventArgs> MessageReceived;

        public void OnMessageReceived()
        {
            var handler = MessageReceived;
            if (handler != null)
            {
                var args = new HookEventArgs(HotkeyInfo.Keys);
                handler.Invoke(this, args);
            }
        }

        public void Attach()
        {
            var hotKey = Key;
            var shift = GetModifierValue(ref hotKey);
            var key = GetKey(hotKey);

            if (key == Keys.None)
            {
                throw new InvalidOperationException(RubberduckUI.CommonHotkey_InvalidKey);
            }

            HookKey(key, shift);
            SetCommandShortcutText();
        }

        public void Detach()
        {
            if (!IsAttached)
            {
                return;
            }

            if (!User32.UnregisterHotKey(_hWndVbe, new IntPtr(HotkeyInfo.HookId)))
            {
                var error = Marshal.GetLastWin32Error();
                Logger.Warn($"Error calling UnregisterHotKey on hokey with id {HotkeyInfo.HookId} for command of type {Command.GetType()}; the error was {error}; going to delete the atom anyway... The memory may leak.");
            }
            Kernel32.SetLastError(Kernel32.ERROR_SUCCESS);
            Kernel32.GlobalDeleteAtom(HotkeyInfo.HookId);
            var lastError = Marshal.GetLastWin32Error();
            if (lastError != Kernel32.ERROR_SUCCESS)
            {
                Logger.Warn($"Error calling DeleteGlobalAtom; the error was {lastError}, the id {HotkeyInfo.HookId} and the type of the associated command {Command.GetType()}.");
            }

            HotkeyInfo = new HotkeyInfo(0, Combo);
            ClearCommandShortcutText();
        }

        private void HookKey(Keys key, uint shift)
        {
            if (IsAttached)
            {
                return;
            }

            var hookId = Kernel32.GlobalAddAtom(Guid.NewGuid().ToString());
            var error = Marshal.GetLastWin32Error();
            if (hookId == 0)
            {               
                Logger.Warn($"Error calling GlobalAddAtom; the error was {error}; aborting the HookKey operation...");    
                return;
            }

            var success = User32.RegisterHotKey(_hWndVbe, new IntPtr(hookId), shift, (uint)key);
            if (!success)
            {
                Logger.Debug(RubberduckUI.CommonHotkey_KeyNotRegistered, key);
                return;
            }

            HotkeyInfo = new HotkeyInfo(hookId, Combo);
            Logger.Trace($"Hotkey for the associated command {Command.GetType()} was registered with id {HotkeyInfo.HookId}");
        }

        private void SetCommandShortcutText()
        {
            if (Command is CommandBase command)
            {
                command.ShortcutText = HotkeyInfo.ToString();
            }
        }

        private void ClearCommandShortcutText()
        {
            if (Command is CommandBase command)
            {
                command.ShortcutText = string.Empty;
            }
        }
        
        private static readonly IDictionary<char,uint> Modifiers = new Dictionary<char, uint>
        {
            { '+', (uint)KeyModifier.SHIFT },
            { '%', (uint)KeyModifier.ALT },
            { '^', (uint)KeyModifier.CONTROL },
        };

        /// <summary>
        /// Gets the <see cref="KeyModifier"/> values out of a key combination.
        /// </summary>
        /// <param name="key">The hotkey string, returned without the modifiers.</param>
        private static uint GetModifierValue(ref string key)
        {
            uint result = 0;

            for (var i = 0; i < 3; i++)
            {
                var firstChar = key[0];
                if (Modifiers.ContainsKey(firstChar))
                {
                    result |= Modifiers[firstChar];
                }
                else
                {
                    // first character isn't a modifier symbol:
                    break;
                }

                // truncate first character for next iteration:
                key = key.Substring(1);
            }

            return result;
        }

        private static Keys GetCombo(string key)
        {
            return GetKey(key.Trim('%', '^', '+')) // will break with special keys, e.g. {f12}
                   | (key.Contains("%") ? Keys.Alt : Keys.None)
                   | (key.Contains("^") ? Keys.Control : Keys.None)
                   | (key.Contains("+") ? Keys.Shift : Keys.None);
        }

        private static Keys GetKey(string keyCode)
        {
            var result = Keys.None;
            switch (keyCode.Substring(0, 1))
            {
                case "{":
                    _keys.TryGetValue(keyCode, out result);
                    break;
                case "~":
                    result = Keys.Return;
                    break;
                case "`":
                    result = Keys.Oemtilde;
                    break;
                default:
                    if (!string.IsNullOrEmpty(keyCode))
                    {
                        result = (Keys)Enum.Parse(typeof(Keys), keyCode);
                    }
                    break;
            }

            return result;
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
    }
}