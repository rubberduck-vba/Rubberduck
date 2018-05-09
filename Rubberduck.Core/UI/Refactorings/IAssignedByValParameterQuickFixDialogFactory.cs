using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.UI.Refactorings
{
    public interface IAssignedByValParameterQuickFixDialogFactory
    {
        IAssignedByValParameterQuickFixDialog Create(string identifier, string identifierType, IEnumerable<string> forbiddenNames);
        void Release(IAssignedByValParameterQuickFixDialog dialog);
    }
}
