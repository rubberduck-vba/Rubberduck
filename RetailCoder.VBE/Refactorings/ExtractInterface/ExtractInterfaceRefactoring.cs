using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<ExtractInterfacePresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private ExtractInterfaceModel _model;

        public ExtractInterfaceRefactoring(IRefactoringPresenterFactory<ExtractInterfacePresenter> factory,
            IActiveCodePaneEditor editor)
        {
            _factory = factory;
            _editor = editor;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();

            if (_model == null) { return; }
        }

        public void Refactor(QualifiedSelection target)
        {
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            Refactor();
        }
    }
}