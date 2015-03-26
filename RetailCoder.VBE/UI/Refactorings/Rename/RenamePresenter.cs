using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.Rename
{
    public class RenamePresenter
    {
        private readonly VBE _vbe;
        private readonly IRenameView _view;
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;

        public RenamePresenter(VBE vbe, IRenameView view, Declarations declarations, QualifiedSelection selection)
        {
            _vbe = vbe;
            _view = view;
            _view.OkButtonClicked += OnOkButtonClicked;

            _declarations = declarations;
            _selection = selection;
        }

        public void Show()
        {
            AcquireTarget(_selection);
            _view.ShowDialog();
        }

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            if (_view.Target.DeclarationType == DeclarationType.Class ||
                _view.Target.DeclarationType == DeclarationType.Module)
            {
                RenameModule();
            }
            else
            {
                RenameDeclaration();
            }

            RenameUsages();
        }

        private void RenameModule()
        {
            try
            {
                var module = _vbe.FindCodeModules(_view.Target.QualifiedName.QualifiedModuleName).Single();
                module.Name = _view.NewName;
            }
            catch (COMException exception)
            {
                MessageBox.Show("Could not rename module.", RubberduckUI.RenameDialog_Caption);
            }
        }

        private void RenameDeclaration()
        {
            var module = _vbe.FindCodeModules(_view.Target.QualifiedName.QualifiedModuleName).First();
            var content = module.get_Lines(_view.Target.Selection.StartLine, 1);
            var newContent = GetReplacementLine(content, _view.Target.IdentifierName, _view.NewName);
            module.ReplaceLine(_view.Target.Selection.StartLine, newContent);
        }

        private void RenameUsages()
        {
            var modules = _view.Target.References.GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                var module = _vbe.FindCodeModules(grouping.Key).First();
                foreach (var line in grouping.GroupBy(reference => reference.Selection.StartLine))
                {
                    var content = module.get_Lines(line.Key, 1);
                    var newContent = GetReplacementLine(content, _view.Target.IdentifierName, _view.NewName);
                    module.ReplaceLine(line.Key, newContent);
                }
            }
        }

        private string GetReplacementLine(string content, string target, string newName)
        {
            // until we figure out how to replace actual tokens,
            // this is going to have to be done the ugly way...

            // what we're trying to avoid here,
            // is to replace all instances of "Foo" in "Foo = FooBar" when target is "Foo".

            var result = ' ' + content;
            if (result.Contains(' ' + target))
            {
                result = result.Replace(' ' + target, ' ' + newName);
            }
            if (result.Contains(target + ' '))
            {
                result = result.Replace(target + ' ', newName + ' ');
            }
            if (result.Contains(target + '.'))
            {
                result = result.Replace(target + '.', newName + '.');
            }
            else if (result.Contains('.' + target))
            {
                result = result.Replace('.' + target, '.'+ newName);
            }

            if (result.Contains(target + '!'))
            {
                result = result.Replace(target + '!', newName + '!');
            }
            else if (result.Contains('!' + target))
            {
                result = result.Replace('!' + target, '!' + newName);
            }

            return result.Substring(1);
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
