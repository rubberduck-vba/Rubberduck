using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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

        private static readonly DeclarationType[] ModuleDeclarationTypes =
            {
                DeclarationType.Class,
                DeclarationType.Module
            };

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            if (ModuleDeclarationTypes.Contains(_view.Target.DeclarationType))
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
                MessageBox.Show(RubberduckUI.RenameDialog_ModuleRenameError, RubberduckUI.RenameDialog_Caption);
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
            return Regex.Replace(content, "\\b" + target + "\\b", newName);
        }

        private static readonly DeclarationType[] ProcedureDeclarationTypes =
            {
                DeclarationType.Procedure,
                DeclarationType.Function,
                DeclarationType.PropertyGet,
                DeclarationType.PropertyLet,
                DeclarationType.PropertySet
            };

        private void AcquireTarget(QualifiedSelection selection)
        {
            _view.Target = _declarations.Items.SingleOrDefault(
                item => IsSelectedDeclaration(selection, item) 
                            || IsSelectedReference(selection, item));

            if (_view.Target == null)
            {
                // rename the containing procedure:
                _view.Target = _declarations.Items.SingleOrDefault(
                    item => ProcedureDeclarationTypes.Contains(item.DeclarationType)
                            && item.Context.GetSelection().Contains(selection.Selection));
            }

            if (_view.Target == null)
            {
                // rename the containing module:
                _view.Target = _declarations.Items.SingleOrDefault(item =>
                    ModuleDeclarationTypes.Contains(item.DeclarationType)
                    && item.QualifiedName.QualifiedModuleName == selection.QualifiedName);
            }
        }

        private bool IsSelectedReference(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.References.Any(r => r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
