using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigations.RegexSearchReplace;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.IdentifierReferences
{
    public class RegexSearchResultsDockablePresenter : DockablePresenterBase
    {
        private static readonly IRubberduckCodePaneFactory CodePaneFactory = new RubberduckCodePaneFactory();

        public RegexSearchResultsDockablePresenter(VBE vbe, AddIn addin, IDockableUserControl control, List<RegexSearchResult> results) : base(vbe, addin, control)
        {
            BindTarget(results);
        }

        private void BindTarget(List<RegexSearchResult> results)
        {
            var listBox = Control.ResultBox;
            listBox.DataSource = results;
            listBox.DisplayMember = "DisplayString";
            Control.Navigate += ControlNavigate;
        }

        public static void OnNavigate(VBE vbe, RegexSearchResult result)
        {
            var project = result.Module.VBE.ActiveVBProject;
            foreach (var proj in result.Module.VBE.VBProjects.Cast<VBProject>().Where(proj => proj.VBComponents.Cast<VBComponent>().Any(v => v.CodeModule == result.Module)))
            {
                project = proj;
                break;
            }
            vbe.SetSelection(project, result.Selection, result.Module.Name, CodePaneFactory);
        }

        private void ControlNavigate(object sender, ListItemActionEventArgs e)
        {
            var reference = e.SelectedItem as RegexSearchResult;
            if (reference != null)
            {
                OnNavigate(VBE, reference);
            }
        }

        SimpleListControl Control { get { return UserControl as SimpleListControl; } }
    }
}