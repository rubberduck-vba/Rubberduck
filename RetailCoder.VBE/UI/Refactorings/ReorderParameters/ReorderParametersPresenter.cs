using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    class ReorderParametersPresenter
    {
        private readonly VBE _vbe;
        private readonly IReorderParametersView _view;
        private readonly QualifiedSelection _selection;
        private readonly VBProjectParseResult _parseResult;

        public ReorderParametersPresenter(VBE vbe, IReorderParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _vbe = vbe;
            _view = view;
            _view.OkButtonClicked += OnOkButtonClicked;

            _parseResult = parseResult;
            _selection = selection;
        }

        public void Show()
        {
            //if (_view.Target != null)
            {
                _view.ShowDialog();
            }
        }

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            
        }

        private void OnCancelButtonClicked(object sender, EventArgs e)
        {

        }
    }
}
