using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<RenamePresenter> _factory;

        public RenameRefactoring(IRefactoringPresenterFactory<RenamePresenter> factory)
        {
            _factory = factory;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            presenter.Show(); //todo: move presenter's renaming logic into this IRefactoring class
        }

        public void Refactor(QualifiedSelection target)
        {
            target.Select();
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            var presenter = _factory.Create();
            presenter.Show(target);
        }
    }
}
