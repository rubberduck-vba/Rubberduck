
using System.Collections.Generic;

namespace Rubberduck.UI.Refactorings
{
    public class AssignedByValParameterQuickFixDialogFactory : IAssignedByValParameterQuickFixDialogFactory
    {
        IAssignedByValParameterQuickFixDialog IAssignedByValParameterQuickFixDialogFactory.Create(string identifier, string identifierType, IEnumerable<string> forbiddenNames)
        {
            return new AssignedByValParameterQuickFixDialog(identifier, identifierType, forbiddenNames);
        }
    }
}
