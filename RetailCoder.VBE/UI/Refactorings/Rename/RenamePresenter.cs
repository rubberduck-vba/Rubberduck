using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
            if (_view.Target != null)
            {
                _view.ShowDialog();
            }
        }

        public void Show(Declaration target)
        {
            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;
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
            catch (COMException)
            {
                MessageBox.Show(RubberduckUI.RenameDialog_ModuleRenameError, RubberduckUI.RenameDialog_Caption);
            }
        }

        private void RenameDeclaration()
        {
            if (_view.Target.DeclarationType == DeclarationType.Control)
            {
                RenameControl();
                return;
            }

            var module = _vbe.FindCodeModules(_view.Target.QualifiedName.QualifiedModuleName).First();
            var content = module.get_Lines(_view.Target.Selection.StartLine, 1);
            var newContent = GetReplacementLine(content, _view.Target.IdentifierName, _view.NewName);
            module.ReplaceLine(_view.Target.Selection.StartLine, newContent);
        }

        private void RenameControl()
        {
            try
            {
                var form = _vbe.FindCodeModules(_view.Target.QualifiedName.QualifiedModuleName).First();
                var control = form.Parent.Designer.Controls(_view.Target.IdentifierName);
                control.Name = _view.NewName;

                foreach (var handler in _declarations.FindEventHandlers(_view.Target))
                {
                    var newMemberName = _view.Target.ComponentName + '_' + _view.NewName;
                    var module = handler.Project.VBComponents.Item(handler.ComponentName).CodeModule;

                    var content = module.get_Lines(handler.Selection.StartLine, 1);
                    var newContent = GetReplacementLine(content, handler.IdentifierName, newMemberName);
                    module.ReplaceLine(handler.Selection.StartLine, newContent);
                }
            }
            catch (COMException)
            {
            }
        }

        private void RenameUsages()
        {
            // todo: refactor

            // rename interface member
            if (_declarations.FindInterfaceMembers().Contains(_view.Target))
            {
                var implementations = _declarations.FindInterfaceImplementationMembers()
                    .Where(m => m.IdentifierName == _view.Target.ComponentName + '_' + _view.Target.IdentifierName);

                foreach (var member in implementations)
                {
                    try
                    {
                        var newMemberName = _view.Target.ComponentName + '_' + _view.NewName;
                        var module = member.Project.VBComponents.Item(member.ComponentName).CodeModule;

                        var content = module.get_Lines(member.Selection.StartLine, 1);
                        var newContent = GetReplacementLine(content, member.IdentifierName, newMemberName);
                        module.ReplaceLine(member.Selection.StartLine, newContent);
                    }
                    catch (COMException)
                    {
                        // gulp
                    }
                }

                return;
            }

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

                // renaming interface
                if (grouping.Any(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext))
                {
                    var members = _declarations.FindMembers(_view.Target);
                    foreach (var member in members)
                    {
                        var oldMemberName = _view.Target.IdentifierName + '_' + member.IdentifierName;
                        var newMemberName = _view.NewName + '_' + member.IdentifierName;
                        var method = _declarations[oldMemberName].SingleOrDefault(m => m.QualifiedName.QualifiedModuleName == grouping.Key);
                        if (method == null)
                        {
                            continue;
                        }

                        var content = module.get_Lines(method.Selection.StartLine, 1);
                        var newContent = GetReplacementLine(content, oldMemberName, newMemberName);
                        module.ReplaceLine(method.Selection.StartLine, newContent);
                    }
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
            var target = _declarations.Items
                .Where(item => item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item) 
                                      || IsSelectedReference(selection, item));

            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;

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

        private void PromptIfTargetImplementsInterface(ref Declaration target)
        {
            var declaration = target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (target == null || interfaceImplementation == null)
            {
                return;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, target.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.RenameDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                target = null;
                return;
            }

            target = interfaceMember;
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
