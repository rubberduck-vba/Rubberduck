using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorRemoveParametersCommand : RefactorCommandBase
    {
        private readonly IMessageBox _msgbox;
        private readonly RubberduckParserState _state;
        private readonly IRewritingManager _rewritingManager;

        public RefactorRemoveParametersCommand(IVBE vbe, RubberduckParserState state, IMessageBox msgbox, IRewritingManager rewritingManager) 
            : base (vbe)
        {
            _msgbox = msgbox;
            _state = state;
            _rewritingManager = rewritingManager;
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

        protected override bool EvaluateCanExecute(object parameter)
        {
            if (_state.Status != ParserState.Ready)
            {
                return false;
            }

            var selection = Vbe.GetActiveSelection();

            if (!selection.HasValue)
            {
                return false;
            }

            var member = _state.AllUserDeclarations.FindTarget(selection.Value, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }
            if (_state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            var parameters = _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .Where(item => member.Equals(item.ParentScopeDeclaration))
                .ToList();
            return member.DeclarationType == DeclarationType.PropertyLet 
                    || member.DeclarationType == DeclarationType.PropertySet
                        ? parameters.Count > 1
                        : parameters.Any();
            
        }

        protected override void OnExecute(object parameter)
        {
            var selection = Vbe.GetActiveSelection();

            using (var view = new RemoveParametersDialog(new RemoveParametersViewModel(_state)))
            {
                var factory = new RemoveParametersPresenterFactory(Vbe, view, _state, _msgbox);
                var refactoring = new RemoveParametersRefactoring(Vbe, factory, _rewritingManager);
                refactoring.Refactor(selection.Value);
            }
        }
    }
}
