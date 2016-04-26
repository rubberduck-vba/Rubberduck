using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodPresenterFactory : IRefactoringPresenterFactory<IExtractMethodPresenter>
    {
        private readonly IActiveCodePaneEditor _editor;
        private readonly IEnumerable<Declaration> _declarations;
        private readonly IIndenter _indenter;

        public ExtractMethodPresenterFactory(IActiveCodePaneEditor editor, IEnumerable<Declaration> declarations, IIndenter indenter)
        {
            _editor = editor;
            _declarations = declarations;
            _indenter = indenter;
        }

        public IExtractMethodPresenter Create()
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
            return new ExtractMethodPresenter(view, model, _indenter);
        }
    }
}