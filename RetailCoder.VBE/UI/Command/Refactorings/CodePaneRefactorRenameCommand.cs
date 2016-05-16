using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public CodePaneRefactorRenameCommand(VBE vbe, RubberduckParserState state, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe)
        {
            _state = state;
            _wrapperWrapperFactory = wrapperWrapperFactory;
        }

        public override bool CanExecute(object parameter)
        {
            if (Vbe.ActiveCodePane == null)
            {
                return false;
            }

            var target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane);
            return _state.Status == ParserState.Ready && target != null && !target.IsBuiltIn;
        }

        public override void Execute(object parameter)
        {
            if (Vbe.ActiveCodePane == null) { return; }

            Declaration target;
            if (parameter != null)
            {
                target = parameter as Declaration;
            }
            else
            {
                target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane);
            }

            if (target == null)
            {
                return;
            }

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state, new MessageBox());
                var refactoring = new RenameRefactoring(Vbe, factory, new MessageBox(), _state);

                refactoring.Refactor(target);
            }
        }
    }
}