using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameModel : IRefactoringModel
    {
        private Declaration _target;
        public Declaration Target
        {
            get => _target;
            set
            {
                _target = value;
                NewName = _target?.IdentifierName ?? string.Empty;
            }
        }

        public Declaration InitialTarget { get; } 
        public bool IsInterfaceMemberRename { set; get; }
        public bool IsControlEventHandlerRename { set; get; }
        public bool IsUserEventHandlerRename { set; get; }

        public string NewName { get; set; }

        public RenameModel(Declaration target)
        {
            Target = target;
            InitialTarget = target;
        }
    }
}
