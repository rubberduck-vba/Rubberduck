using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
