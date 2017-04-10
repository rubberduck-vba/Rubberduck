using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactorProperty : RenameRefactorBase
    {
        public RenameRefactorProperty(RubberduckParserState state)
            : base(state) { _errorMessage = string.Empty; }

        private string _errorMessage;
        override public string ErrorMessage { get { return _errorMessage; } }

        override public void Rename(Declaration renameTarget, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_PropertyRenameError, renameTarget.IdentifierName); ;

            var members = State.AllUserDeclarations.Named(renameTarget.IdentifierName)
                .Where(item => item.ProjectId == renameTarget.ProjectId
                    && item.ComponentName == renameTarget.ComponentName
                    && item.DeclarationType.HasFlag(DeclarationType.Property));


            foreach (var member in members)
            {
                RenameUsages(member, newName);
            }

            RenameDeclaration(renameTarget, newName);

            Rewrite();
        }
    }
}
