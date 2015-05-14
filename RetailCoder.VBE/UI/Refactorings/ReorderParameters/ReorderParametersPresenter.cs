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
        private readonly Declarations _declarations;
        private readonly QualifiedSelection _selection;
        private readonly VBProjectParseResult _parseResult;

        public ReorderParametersPresenter(VBE vbe, IReorderParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _vbe = vbe;
            _view = view;
            _view.OkButtonClicked += OnOkButtonClicked;

            _parseResult = parseResult;
            _declarations = parseResult.Declarations;
            _selection = selection;
        }

        public void Show()
        {
            AcquireTarget(_selection);
            if (_view.Target != null &&
                (
                 _view.Target.DeclarationType == DeclarationType.Event ||
                 _view.Target.DeclarationType == DeclarationType.Function ||
                 _view.Target.DeclarationType == DeclarationType.Procedure ||
                 _view.Target.DeclarationType == DeclarationType.PropertyGet ||
                 _view.Target.DeclarationType == DeclarationType.PropertyLet ||
                 _view.Target.DeclarationType == DeclarationType.PropertySet ||
                 _view.Target.DeclarationType == DeclarationType.LibraryFunction ||
                 _view.Target.DeclarationType == DeclarationType.LibraryProcedure))
            {
                _view.ShowDialog();
            }
        }

        private void OnOkButtonClicked(object sender, EventArgs e)
        {
            
        }

        private static readonly DeclarationType[] ModuleDeclarationTypes =
            {
                DeclarationType.Class,
                DeclarationType.Module
            };

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
                .Where(item => !item.IsBuiltIn && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                      || IsSelectedReference(selection, item));

            PromptIfTargetImplementsInterface(ref target);
            _view.Target = target;

            if (_view.Target == null)
            {
                return;

                // rename the containing procedure:
                _view.Target = _declarations.Items.SingleOrDefault(
                    item => !item.IsBuiltIn
                            && ProcedureDeclarationTypes.Contains(item.DeclarationType)
                            && item.Context.GetSelection().Contains(selection.Selection));
            }

            if (_view.Target == null)
            {
                return;
                // rename the containing module:
                _view.Target = _declarations.Items.SingleOrDefault(item =>
                    !item.IsBuiltIn
                    && ModuleDeclarationTypes.Contains(item.DeclarationType)
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
            return declaration.References.Any(r =>
                r.QualifiedModuleName == selection.QualifiedName &&
                r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
