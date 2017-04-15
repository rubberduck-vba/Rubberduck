using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactorEvent : RenameRefactorBase
    {
        public RenameRefactorEvent(RubberduckParserState state)
            : base(state) { _errorMessage = string.Empty; }

        private string _errorMessage;
        public override string ErrorMessage => _errorMessage;

        public override void Rename(Declaration eventDeclaration, string newName)
        {
            _errorMessage = string.Format(RubberduckUI.RenameDialog_EventRenameError, eventDeclaration.IdentifierName);
            Debug.Assert(eventDeclaration.DeclarationType == DeclarationType.Event, "Resolving User Selection to Declaration Failed");

            RenameUsages(eventDeclaration, newName);

            var handlers = State.AllUserDeclarations.FindHandlersForEvent(eventDeclaration).ToList();
            handlers.ForEach(handler => RenameDeclaration(handler.Item2, handler.Item1.IdentifierName + '_' + newName));

            RenameDeclaration(eventDeclaration, newName);

            Rewrite();
        }
    }
}
