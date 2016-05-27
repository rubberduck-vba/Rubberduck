using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRemoveParametersCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorRemoveParametersCommand(VBE vbe, RubberduckParserState state) 
            : base (vbe)
        {
            _state = state;
        }

        private static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
        
        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null || _state.Status != ParserState.Ready)
            {
                return false;
            }

            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            var member = _state.AllUserDeclarations.FindTarget(selection.Value, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }

            var parameters = _state.AllUserDeclarations.Where(item => item.DeclarationType == DeclarationType.Parameter && member.Equals(item.ParentScopeDeclaration)).ToList();
            var canExecute = (member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet)
                    ? parameters.Count > 1
                    : parameters.Any();

            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            var selection = Vbe.ActiveCodePane.GetQualifiedSelection();
            using (var view = new RemoveParametersDialog())
            {
                var factory = new RemoveParametersPresenterFactory(Vbe, view, _state, new MessageBox());
                var refactoring = new RemoveParametersRefactoring(Vbe, factory);
                refactoring.Refactor(selection.Value);
            }
        }
    }
}
