using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using System.Runtime.InteropServices;
using NLog.Targets;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Settings;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        public CodePaneRefactorRenameCommand(VBE vbe, RubberduckParserState state, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperWrapperFactory) 
            : base (vbe, editor)
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
            return _state.Status == ParserState.Ready && target != null;
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
                var selection = Vbe.ActiveCodePane.GetSelection();
                target = _state.AllUserDeclarations.FindTarget(selection);
            }

            if (target == null)
            {
                return;
            }

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state, new MessageBox(), _wrapperWrapperFactory);
                var refactoring = new RenameRefactoring(factory, Editor, new MessageBox(), _state);

                refactoring.Refactor(target);
            }
        }
    }
}