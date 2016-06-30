using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Input;
using NLog;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public abstract class CommandBase : ICommand
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public virtual bool CanExecuteImpl(object parameter)
        {
            return true;
        }

        public abstract void ExecuteImpl(object parameter);

        public bool CanExecute(object parameter)
        {
            try
            {
                return CanExecuteImpl(parameter);
            }
            catch (Exception e)
            {
                Logger.Fatal(e);

                System.Windows.Forms.MessageBox.Show(
                    RubberduckUI.RubberduckFatalError, RubberduckUI.Rubberduck,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }
        }

        public void Execute(object parameter)
        {
            try
            {
                ExecuteImpl(parameter);
            }
            catch (Exception e)
            {
                Logger.Fatal(e);

                System.Windows.Forms.MessageBox.Show(
                    RubberduckUI.RubberduckFatalError, RubberduckUI.Rubberduck,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public virtual string ShortcutText { get; set; }
        
        public virtual RubberduckHotkey Hotkey { get { return RubberduckHotkey.None; } }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
