using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    class RenameRefactorInterface : RenameRefactorBase
    {
        public RenameRefactorInterface(RubberduckParserState state)
            : base(state) { _errorMessage = string.Empty; }

        private string _errorMessage;
        public override string ErrorMessage => _errorMessage;

        public override void Rename(Declaration renameTarget, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_InterfaceRenameError, renameTarget.IdentifierName);

            RenameUsages(renameTarget, newName);

            var implementations = State.AllUserDeclarations.FindInterfaceImplementationMembers()
                .Where(m => m.IdentifierName == renameTarget.ComponentName + '_' + renameTarget.IdentifierName);

            var newMemberName = renameTarget.ComponentName + '_' + newName;
            foreach (var imp in implementations)
            {
                RenameDeclaration(imp, newMemberName);
            }

            RenameDeclaration(renameTarget, newName);

            Rewrite();
        }
    }
}
