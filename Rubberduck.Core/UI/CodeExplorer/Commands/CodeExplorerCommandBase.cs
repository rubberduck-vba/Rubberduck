using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public abstract class CodeExplorerCommandBase : ComCommandBase
    {
        protected CodeExplorerCommandBase(IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        public abstract IEnumerable<Type> ApplicableNodeTypes { get; }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return parameter != null && ApplicableNodeTypes.Contains(parameter.GetType());
        }
    }
}
