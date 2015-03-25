using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.Rename
{
    public class RenamePresenter
    {
        private readonly IRenameView _view;
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;

        public RenamePresenter(IRenameView view, Declarations declarations, QualifiedSelection selection)
        {
            _view = view;
            _view.OkButtonClicked += OnOkButtonClicked;

            _declarations = declarations;
            _selection = selection;
        }

        public void Show()
        {
            AcquireTarget(_selection);
            if (_view.Target == null)
            {
                return; // something's wrong. should we tell 'em?
            }

            _view.ShowDialog();
        }

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            var targetSelection = new QualifiedSelection(_view.Target.QualifiedName.QualifiedModuleName, _view.Target.Selection);
            var usages = _view.Target.References;

            MessageBox.Show("renaming " + targetSelection + " to '" + _view.NewName + "'...");
        }

        private void AcquireTarget(QualifiedSelection selection)
        {
            var targets = _declarations.Items.Where(declaration =>
                declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                && (declaration.Selection.Contains(selection.Selection))
                || declaration.References.Any(r => r.Selection.Contains(selection.Selection)))
                .ToList();

            var nonProcTarget = targets.Where(t => t.DeclarationType != DeclarationType.Function
                                                   && t.DeclarationType != DeclarationType.Procedure
                                                   && t.DeclarationType != DeclarationType.PropertyGet
                                                   && t.DeclarationType != DeclarationType.PropertyLet
                                                   && t.DeclarationType != DeclarationType.PropertySet).ToList();
            if (nonProcTarget.Any())
            {
                _view.Target = nonProcTarget.First();
            }
            else
            {
                _view.Target = targets.FirstOrDefault();
            }

            if (_view.Target == null)
            {
                // no valid selection? no problem - let's rename the module:
                _view.Target = _declarations.Items.SingleOrDefault(declaration =>
                    declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                    && declaration.DeclarationType == DeclarationType.Class ||
                    declaration.DeclarationType == DeclarationType.Module);
            }
        }
    }
}
