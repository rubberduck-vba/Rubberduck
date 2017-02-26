
namespace Rubberduck.UI.Refactorings
{
    public class AssignedByValParameterQuickFixMockDialogFactory : IAssignedByValParameterQuickFixDialogFactory
    {
        private string _userEnteredVariableName;
        public AssignedByValParameterQuickFixMockDialogFactory()
        {
            _userEnteredVariableName = string.Empty;
        }
        public AssignedByValParameterQuickFixMockDialogFactory(string userEnteredVariableName)
        {
            _userEnteredVariableName = userEnteredVariableName;
        }
        IAssignedByValParameterQuickFixDialog IAssignedByValParameterQuickFixDialogFactory.Create(string identifier, string identifierType)
        {
            return new AssignedByValParameterQuickFixMockDialog(_userEnteredVariableName);
        }
    }
}
