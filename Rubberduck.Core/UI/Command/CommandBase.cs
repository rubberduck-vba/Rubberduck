using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Input;
using NLog;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public abstract class CommandBase : ICommand
    {
        private static readonly List<MethodBase> ExceptionTargetSites = new List<MethodBase>();

        protected CommandBase(ILogger logger = null)
        {
            Logger = logger ?? LogManager.GetLogger(GetType().FullName);
            CanExecuteCondition = (parameter => true);
        }

        protected ILogger Logger { get; }
        protected abstract void OnExecute(object parameter);

        protected Func<object, bool> CanExecuteCondition { get; private set; }

        protected void AddToCanExecuteEvaluation(Func<object, bool> furtherCanExecuteEvaluation)
        {
            if (furtherCanExecuteEvaluation == null)
            {
                return;
            }

            var currentCanExecute = CanExecuteCondition;
            CanExecuteCondition = (parameter) => currentCanExecute(parameter) && furtherCanExecuteEvaluation(parameter);
        }

        public bool CanExecute(object parameter)
        {
            try
            {
                return CanExecuteCondition(parameter);
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

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
