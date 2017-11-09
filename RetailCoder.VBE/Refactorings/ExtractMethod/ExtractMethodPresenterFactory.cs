using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodPresenterFactory : IRefactoringPresenterFactory<IExtractMethodPresenter>
    {
        private readonly IIndenter _indenter;
        private readonly RubberduckParserState _state;
        private readonly QualifiedSelection _selection;

        public ExtractMethodPresenterFactory(IIndenter indenter, RubberduckParserState state, QualifiedSelection selection)
        {
            _indenter = indenter;
            _state = state;
            _selection = selection;
        }

        public IExtractMethodPresenter Create()
        {
            var view = new ExtractMethodDialog(new ExtractMethodViewModel());
            var model = new ExtractMethodModel(_state, _selection);
            return new ExtractMethodPresenter(view, model, _indenter);
        }
    }
}
