using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    class RefactorReorderParametersCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public RefactorReorderParametersCommand(VBE vbe, RubberduckParserState state, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe)
        {
            _state = state;
            _wrapperWrapperFactory = wrapperWrapperFactory;
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
            var member = _state.AllUserDeclarations.FindTarget(selection, ValidDeclarationTypes);
            if (member == null)
            {
                return false;
            }

            var parameters = _state.AllUserDeclarations.Where(item => member.Equals(item.ParentScopeDeclaration)).ToList();
            var canExecute = (member.DeclarationType == DeclarationType.PropertyLet || member.DeclarationType == DeclarationType.PropertySet)
                    ? parameters.Count > 2
                    : parameters.Count > 1;

            Debug.WriteLine("{0}.CanExecute evaluates to {1}", GetType().Name, canExecute);
            return canExecute;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return;
            }
            var codePane = _wrapperWrapperFactory.Create(Vbe.ActiveCodePane);
            var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), codePane.Selection);

            using (var view = new ReorderParametersDialog())
            {
                var factory = new ReorderParametersPresenterFactory(Vbe, view, _state, new MessageBox());
                var refactoring = new ReorderParametersRefactoring(Vbe, factory, new MessageBox());
                refactoring.Refactor(selection);
            }
        }
    }
}
