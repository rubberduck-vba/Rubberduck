using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactoring.ExtractMethod
{
    public class ExtractMethodPresenterFactory : IRefactoringPresenterFactory<ExtractMethodPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly Declarations _declarations;

        public ExtractMethodPresenterFactory(IActiveCodePaneEditor editor, Declarations declarations)
        {
            _editor = editor;
            _declarations = declarations;
        }

        public ExtractMethodPresenter Create()
        {
            var selection = _editor.GetSelection();
            if (selection == null)
            {
                return null;
            }

            ExtractMethodModel model;
            try
            {
                model = new ExtractMethodModel(_editor, _declarations, selection.Value);
            }
            catch (InvalidOperationException)
            {
                return null;
            }

            var view = new ExtractMethodDialog();
            return new ExtractMethodPresenter(view, model);
        }
    }
}