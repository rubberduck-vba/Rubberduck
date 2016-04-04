using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRemoveParametersCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;

        public RefactorRemoveParametersCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor) 
            : base (vbe, editor)
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

            var selection = Vbe.ActiveCodePane.GetSelection();
            var member = _state.AllUserDeclarations.FindTarget(selection, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }

            var parameters = _state.AllUserDeclarations.Where(item => member.Equals(item.ParentScopeDeclaration)).ToList();
            var canExecute = (member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet)
                    ? parameters.Count > 1
                    : parameters.Any();

            Debug.WriteLine("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }

            var selection = Vbe.ActiveCodePane.GetSelection();
            using (var view = new RemoveParametersDialog())
            {
                var factory = new RemoveParametersPresenterFactory(Editor, view, _state, new MessageBox());
                var refactoring = new RemoveParametersRefactoring(factory, Editor);
                refactoring.Refactor(selection);
            }
        }
    }
}