using System;
using NLog;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    public abstract class ComCommandBase : CommandBase
    {
        private readonly IVbeEvents _vbeEvents;

        protected ComCommandBase(ILogger logger, IVbeEvents vbeEvents) : base(logger)
        {
            _vbeEvents = vbeEvents;
        }
        
        protected override bool EvaluateCanExecute(object parameter)
        {
            return !_vbeEvents.Terminated && base.EvaluateCanExecute(parameter);
        }

        public new void Execute(object parameter)
        {
            if (_vbeEvents.Terminated)
            {
                return;
            }

            base.Execute(parameter);
        }
    }
}
