using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule;

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
            return new ExtractMethodPresenter(view, _indenter);
        }
    }
}