using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    class RenameRefactorPropertyParameter : RenameRefactorBase
    {
        public RenameRefactorPropertyParameter(RubberduckParserState state)
            : base(state) { _errorMessage = string.Empty; }

        private string _errorMessage;
        public override string ErrorMessage => _errorMessage;

        public override void Rename(Declaration renameTarget, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_PropertyParameterRenameError, renameTarget.IdentifierName, renameTarget.ParentDeclaration.IdentifierName);

            var parameters = State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Where(d =>
                d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
                && d.IdentifierName == renameTarget.IdentifierName);

            foreach (var param in parameters)
            {
                RenameUsages(param, newName);
            }

            foreach (var param in parameters)
            {
                RenameDeclaration(param, newName);
            }

            Rewrite();
        }
    }
}