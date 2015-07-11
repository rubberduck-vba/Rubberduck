using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenterFactory : IRefactoringPresenterFactory<RenamePresenter>
    {
        private readonly VBE _vbe;
        private readonly IRenameView _view;
        private readonly VBProjectParseResult _parseResult;
        private readonly IMessageBox _messageBox;
        private readonly IRubberduckFactory<IRubberduckCodePane> _factory;

        public RenamePresenterFactory(VBE vbe, IRenameView view, VBProjectParseResult parseResult, IMessageBox messageBox, IRubberduckFactory<IRubberduckCodePane> factory)
        {
            _vbe = vbe;
            _view = view;
            _parseResult = parseResult;
            _messageBox = messageBox;
            _factory = factory;
        }

        public RenamePresenter Create()
        {
            var codePane = _factory.Create(_vbe.ActiveCodePane);
            var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent),
                codePane.Selection, _factory);
            return new RenamePresenter(_view, new RenameModel(_vbe, _parseResult, selection, _messageBox));
        }
    }
}