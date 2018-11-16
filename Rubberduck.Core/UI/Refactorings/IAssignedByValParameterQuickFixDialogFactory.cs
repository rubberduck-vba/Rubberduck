using System;

namespace Rubberduck.UI.Refactorings
{
    public interface IAssignedByValParameterQuickFixDialogFactory
    {
        IAssignedByValParameterQuickFixDialog Create(string identifier, string identifierType, Func<string, bool> nameCollisionChecker);
        void Release(IAssignedByValParameterQuickFixDialog dialog);
    }
}
