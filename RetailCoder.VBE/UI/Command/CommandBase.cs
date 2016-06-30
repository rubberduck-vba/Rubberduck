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
        protected CommandBase(ILogger logger)
        {
            _logger = logger;
        }

        private readonly ILogger _logger;
        protected virtual ILogger Logger { get { return _logger; } }

        protected virtual bool CanExecuteImpl(object parameter)
        {
            return true;
        }

        protected abstract void ExecuteImpl(object parameter);

        public bool CanExecute(object parameter)
        {
            try
            {
                return CanExecuteImpl(parameter);
            }
            catch (Exception exception)
            {
                _logger.Fatal(exception);

                var messageBox = new MessageBox();
                messageBox.Show(
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
            catch (Exception exception)
            {
                _logger.Fatal(exception);

                var messageBox = new MessageBox();
                messageBox.Show(
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
