using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodPresenterFactory : IRefactoringPresenterFactory<IExtractMethodPresenter>
    {
        private readonly IIndenter _indenter;

        public ExtractMethodPresenterFactory(IIndenter indenter)
        {
            _indenter = indenter;
        }

        public IExtractMethodPresenter Create()
        {
            var view = new ExtractMethodDialog();
            var model = new ExtractMethodModel();
            return new ExtractMethodPresenter(view, model, _indenter);
        }
    }
}
