using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactorDefault : RenameRefactorBase
    {
        public RenameRefactorDefault(RubberduckParserState state)
            : base(state) { _errorMessage = string.Empty; }

        private string _errorMessage;
        public override string ErrorMessage => _errorMessage;

        public override void Rename(Declaration renameTarget, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_DefaultRenameError, renameTarget.IdentifierName);

            RenameUsages(renameTarget, newName);

            RenameDeclaration(renameTarget, newName);

            Rewrite();
        }
    }
}
