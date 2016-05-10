using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenterFactory : IRefactoringPresenterFactory<RenamePresenter>
    {
        private readonly VBE _vbe;
        private readonly IRenameDialog _view;
        private readonly RubberduckParserState _parseResult;
        private readonly IMessageBox _messageBox;

        public RenamePresenterFactory(VBE vbe, IRenameDialog view, RubberduckParserState parseResult, IMessageBox messageBox)
        {
            _vbe = vbe;
            _view = view;
            _parseResult = parseResult;
            _messageBox = messageBox;
        }

        public RenamePresenter Create()
        {
            var codePane = _vbe.ActiveCodePane;
            var selection = codePane.GetSelection();
            
            var qualifiedSelection = codePane == null || selection == null 
                ? new QualifiedSelection() 
                : new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), selection.Value);

            return new RenamePresenter(_view, new RenameModel(_vbe, _parseResult, qualifiedSelection, _messageBox));
        }
    }
}