using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings
{
    public class AssignedByValParameterQuickFixDialogFactory : IAssignedByValParameterQuickFixDialogFactory
    {
        IAssignedByValParameterQuickFixDialog IAssignedByValParameterQuickFixDialogFactory.Create(string identifier, string identifierType)
        {
            return new AssignedByValParameterQuickFixDialog(identifier, identifierType);
        }
    }
}
