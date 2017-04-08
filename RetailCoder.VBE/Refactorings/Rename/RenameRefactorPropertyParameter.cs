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
        override public string ErrorMessage { get { return _errorMessage; } }

        override public void Rename(Declaration renameTarget, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_PropertyParameterRenameError, renameTarget.IdentifierName, renameTarget.ParentDeclaration.IdentifierName);

            var parameters = State.AllUserDeclarations.Where(d =>
                d.DeclarationType == DeclarationType.Parameter
                && d.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Property)
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