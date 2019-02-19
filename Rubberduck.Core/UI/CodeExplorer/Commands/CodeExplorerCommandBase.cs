using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public abstract class CodeExplorerCommandBase : CommandBase
    {
        protected CodeExplorerCommandBase() : base(LogManager.GetCurrentClassLogger()) { }

        public abstract IEnumerable<Type> ApplicableNodeTypes { get; }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return parameter != null && ApplicableNodeTypes.Contains(parameter.GetType());
        }
    }
}
