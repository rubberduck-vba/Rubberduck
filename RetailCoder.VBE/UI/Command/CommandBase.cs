using System;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public abstract class CommandBase : ICommand
    {
        public virtual bool CanExecute(object parameter)
        {
            return true;
        }

        public abstract void Execute(object parameter);

        public virtual string ShortcutText { get; set; }
        
        public virtual RubberduckHotkey Hotkey { get { return RubberduckHotkey.None; } }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
