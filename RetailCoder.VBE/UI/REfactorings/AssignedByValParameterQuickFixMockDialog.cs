using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings
{
    public class AssignedByValParameterQuickFixMockDialog : IAssignedByValParameterQuickFixDialog
    {
        private string _testLocalVariableName;
        internal AssignedByValParameterQuickFixMockDialog(string testLocalVariableName = "")
        {
            _testLocalVariableName = testLocalVariableName;
        }
        public DialogResult ShowDialog() { return DialogResult.OK; }

        public void Dispose()
        {
        }
        public DialogResult DialogResult { set; get; }
        private string _newName;
        public string NewName
        {
            get
            {
                if (_testLocalVariableName.Length > 0)
                {
                    return _testLocalVariableName;
                }
                else
                {
                    return _newName;
                }
            }
            set { _newName = value; }
        }
        public string[] IdentifierNamesAlreadyDeclared { get; set; }
    }
}
