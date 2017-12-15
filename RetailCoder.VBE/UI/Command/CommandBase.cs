using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Input;
using NLog;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public abstract class CommandBase : ICommand
    {
        private static readonly List<MethodBase> ExceptionTargetSites = new List<MethodBase>();

        protected CommandBase(ILogger logger)
        {
            Logger = logger;
        }

        protected ILogger Logger { get; }
        protected virtual bool EvaluateCanExecute(object parameter) => true;
        protected abstract void OnExecute(object parameter);

        public bool CanExecute(object parameter)
        {
            try
            {
                return EvaluateCanExecute(parameter);
            }
            catch (Exception exception)
            {
                Logger.Error(exception);

                if (!ExceptionTargetSites.Contains(exception.TargetSite))
                {
                    ExceptionTargetSites.Add(exception.TargetSite);
                }

                return false;
            }
        }

        public void Execute(object parameter)
        {
            try
            {
                OnExecute(parameter);
            }
            catch (Exception exception)
            {
                Logger.Error(exception);

                if (!ExceptionTargetSites.Contains(exception.TargetSite))
                {
                    ExceptionTargetSites.Add(exception.TargetSite);
                }
            }
        }

        public string ShortcutText { get; set; }
        
        public virtual RubberduckHotkey Hotkey => RubberduckHotkey.None;

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
