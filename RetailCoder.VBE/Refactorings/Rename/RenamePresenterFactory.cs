using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenterFactory : IRefactoringPresenterFactory<RenamePresenter>
    {
        private readonly VBE _vbe;
        private readonly IRenameView _view;
        private readonly IRubberduckParserState _parseResult;
        private readonly IMessageBox _messageBox;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public RenamePresenterFactory(VBE vbe, IRenameView view, IRubberduckParserState parseResult, IMessageBox messageBox, ICodePaneWrapperFactory wrapperFactory)
        {
            _vbe = vbe;
            _view = view;
            _parseResult = parseResult;
            _messageBox = messageBox;
            _wrapperFactory = wrapperFactory;
        }

        public RenamePresenter Create()
        {
            var codePane = _wrapperFactory.Create(_vbe.ActiveCodePane);
            var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent),
                codePane.Selection);
            return new RenamePresenter(_view, new RenameModel(_vbe, _parseResult, selection, _messageBox));
        }
    }
}