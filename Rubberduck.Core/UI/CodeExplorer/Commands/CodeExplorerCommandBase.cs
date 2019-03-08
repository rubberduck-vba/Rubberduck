using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public abstract class CodeExplorerCommandBase : ComCommandBase
    {
        protected CodeExplorerCommandBase(IVBEEvents vbeEvents) : base(LogManager.GetCurrentClassLogger(), vbeEvents) { }

        public abstract IEnumerable<Type> ApplicableNodeTypes { get; }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return parameter != null && ApplicableNodeTypes.Contains(parameter.GetType());
        }
    }
}
