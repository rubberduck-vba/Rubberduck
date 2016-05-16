using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodPresenterFactory : IRefactoringPresenterFactory<IExtractMethodPresenter>
    {
        private readonly VBE _vbe;
        private readonly IEnumerable<Declaration> _declarations;
        private readonly IIndenter _indenter;

        public ExtractMethodPresenterFactory(VBE vbe, IEnumerable<Declaration> declarations, IIndenter indenter)
        {
            _vbe = vbe;
            _declarations = declarations;
            _indenter = indenter;
        }

        public IExtractMethodPresenter Create()
        {
            var selection = _vbe.ActiveCodePane.CodeModule.GetSelection();
            if (selection == null)
            {
                return null;
            }

            ExtractMethodModel model;
            try
            {
                model = new ExtractMethodModel(_vbe, _declarations, selection.Value);
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