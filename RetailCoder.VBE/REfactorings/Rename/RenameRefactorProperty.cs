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
        public override string ErrorMessage => _errorMessage;

        public override void Rename(Declaration renameTarget, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_PropertyRenameError, renameTarget.IdentifierName); ;

            var members = State.DeclarationFinder.MatchName(renameTarget.IdentifierName)
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
