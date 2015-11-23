using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.IdentifierReferences;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Navigation
{
    public class NavigateAllReferences : IDeclarationNavigator
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly IRubberduckParser _parser;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly IMessageBox _messageBox;

        public NavigateAllReferences(VBE vbe, AddIn addin, IRubberduckParser parser, ICodePaneWrapperFactory wrapperFactory, IMessageBox messageBox)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;
            _wrapperFactory = wrapperFactory;
            _messageBox = messageBox;
        }

        public void Find()
        {
            //var codePane = _wrapperFactory.Create(_vbe.ActiveCodePane);
            //var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), codePane.Selection);
            //var progress = new ParsingProgressPresenter();
            //var result = progress.Parse(_parser, _vbe.ActiveVBProject);
            //if (result == null)
            //{
            //    return; // bug/todo: something's definitely wrong, exception thrown in resolver code
            //}

            //var declarations = result.Declarations.Items
            //    .Where(item => item.DeclarationType != DeclarationType.ModuleOption)
            //    .ToList();

            //var target = declarations.SingleOrDefault(item =>
            //    item.IsSelected(selection)
            //    || item.References.Any(r => r.IsSelected(selection)));

            //if (target != null)
            //{
            //    Find(target);
            //}
        }

        public void Find(Declaration target)
        {
            var referenceCount = target.References.Count();

            if (referenceCount == 1)
            {
                // if there's only 1 reference, just jump to it:
                IdentifierReferencesListDockablePresenter.OnNavigateIdentifierReference(_vbe, target.References.First());
            }
            else if (referenceCount > 1)
            {
                // if there's more than one reference, show the dockable reference navigation window:
                try
                {
                    ShowReferencesToolwindow(target);
                }
                catch (COMException)
                {
                    // the exception is related to the docked control host instance,
                    // trying again will work (I know, that's bad bad bad code)
                    ShowReferencesToolwindow(target);
                }
            }
            else
            {
                var message = string.Format(RubberduckUI.AllReferences_NoneFound, target.IdentifierName);
                var caption = string.Format(RubberduckUI.AllReferences_Caption, target.IdentifierName);
                _messageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowReferencesToolwindow(Declaration target)
        {
            // throws a COMException if toolwindow was already closed
            var window = new SimpleListControl(target);
            var presenter = new IdentifierReferencesListDockablePresenter(_vbe, _addin, window, target, _wrapperFactory);
            presenter.Show();
        }
    }
}