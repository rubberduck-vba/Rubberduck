using System;
using NLog;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    public abstract class ComCommandBase : CommandBase
    {
        private IVBEEvents _vbeEvents;
        private bool _terminated;

        protected ComCommandBase(ILogger logger, IVBEEvents vbeEvents) : base(logger)
        {
            _vbeEvents = vbeEvents;
            _vbeEvents.EventsTerminated += HandleEventsTerminated;
        }

        private void HandleEventsTerminated(object sender, EventArgs e)
        {
            _terminated = true;
            _vbeEvents = null;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return !_terminated && base.EvaluateCanExecute(parameter);
        }

        public new void Execute(object parameter)
        {
            if (_terminated)
            {
                return;
            }

            base.Execute(parameter);
        }
    }
}
