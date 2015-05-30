using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBA;
using Rubberduck.VBEditor;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    class RemoveParametersPresenter
    {
        private readonly IRemoveParametersView _view;
        private readonly Declarations _declarations;

        public RemoveParametersPresenter(IRemoveParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _view = view;
            _view.RemoveParams = new Refactoring.RemoveParameterRefactoring.RemoveParameterRefactoring(parseResult, selection);

            _declarations = parseResult.Declarations;

            _view.OkButtonClicked += OkButtonClicked;
        }

        public void Show()
        {
            _view.InitializeParameterGrid();
            _view.ShowDialog();
        }

        private void OkButtonClicked(object sender, EventArgs e)
        {
            if (!_view.RemoveParams.Parameters.Where(item => item.IsRemoved).Any())
            {
                return;
            }

            _view.RemoveParams.Refactor();
        }
    }
}
