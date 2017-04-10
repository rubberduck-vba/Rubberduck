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
        override public string ErrorMessage { get { return _errorMessage; } }

        override public void Rename(Declaration renameTarget, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_DefaultRenameError, renameTarget.IdentifierName);

            RenameUsages(renameTarget, newName);

            RenameDeclaration(renameTarget, newName);

            Rewrite();
        }
    }
}
