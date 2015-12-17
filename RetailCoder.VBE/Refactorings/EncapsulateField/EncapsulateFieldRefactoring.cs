using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    class EncapsulateFieldRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<IEncapsulateFieldPresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private EncapsulateFieldModel _model;

        public EncapsulateFieldRefactoring(IRefactoringPresenterFactory<IEncapsulateFieldPresenter> factory, IActiveCodePaneEditor editor)
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

            Refactor(_model.TargetDeclaration);
        }

        public void Refactor(QualifiedSelection target)
        {
        }

        public void Refactor(Declaration target)
        {
        }
    }
}
