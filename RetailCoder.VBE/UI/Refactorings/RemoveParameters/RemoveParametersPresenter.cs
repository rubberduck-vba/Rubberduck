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
        }

        public void Show()
        {
            _view.InitializeParameterGrid();
            _view.ShowDialog();
        }

        private void PromptIfTargetImplementsInterface(ref Declaration target, ref Declaration method)
        {
            /*var declaration = method;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (method == null || interfaceImplementation == null)
            {
                return;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.ReorderPresenter_TargetIsInterfaceMemberImplementation, method.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                method = null;
                return;
            }

            method = interfaceMember;

            var proc = (dynamic)declaration.Context;
            var paramList = (VBAParser.ArgListContext)proc.argList();

            var indexOfInterfaceParam = paramList.arg().ToList().FindIndex(item => item.GetText() == _target.Context.GetText());
            target = FindTargets(_method).ElementAt(indexOfInterfaceParam);*/
        }
    }
}
