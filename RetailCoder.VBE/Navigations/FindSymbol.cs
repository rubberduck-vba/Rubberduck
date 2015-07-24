using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.UI;
using Rubberduck.UI.FindSymbol;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Navigations
{
    public class FindSymbol : IFindSymbol
    {
        private readonly VBE _vbe;
        private readonly IRubberduckParser _parser;
        private readonly SearchResultIconCache _iconCache;
        private readonly IRubberduckCodePaneFactory _codePaneFactory;

        public FindSymbol(VBE vbe, IRubberduckParser parser, IRubberduckCodePaneFactory codePaneFactory)
        {
            _vbe = vbe;
            _parser = parser;
            _iconCache = new SearchResultIconCache();
            _codePaneFactory = codePaneFactory;
        }

        public void Find()
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, _vbe.ActiveVBProject);
            var declarations = result.Declarations;
            var vm = new FindSymbolViewModel(declarations.Items.Where(item => !item.IsBuiltIn), _iconCache);
            using (var view = new FindSymbolDialog(vm))
            {
                view.Navigate += view_Navigate;
                view.ShowDialog();
            }
        }

        private void view_Navigate(object sender, NavigateCodeEventArgs e)
        {
            if (e.QualifiedName.Component == null)
            {
                return;
            }

            try
            {
                var codePane = _codePaneFactory.Create(e.QualifiedName.Component.CodeModule.CodePane);
                codePane.Selection = e.Selection;
            }
            catch (COMException)
            {
            }
        }
    }
}