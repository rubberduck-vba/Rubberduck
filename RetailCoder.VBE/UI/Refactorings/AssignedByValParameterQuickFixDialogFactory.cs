
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
