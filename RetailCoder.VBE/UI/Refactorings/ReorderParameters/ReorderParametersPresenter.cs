using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactoring.ReorderParametersRefactoring;
using Rubberduck.VBEditor;
using System;
using System.Linq;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    class ReorderParametersPresenter
    {
        private readonly IReorderParametersView _view;
        private readonly Declarations _declarations;
        
        public ReorderParametersPresenter(IReorderParametersView view, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _view = view;
            _declarations = parseResult.Declarations;

            _view.ReorderParams = new ReorderParametersRefactoring(parseResult, selection);
            _view.ReorderParams.Target = PromptIfTargetImplementsInterface();

            _view.OkButtonClicked += OkButtonClicked;
        }

        /// <summary>
        /// Displays the Refactor Parameters dialog window.
        /// </summary>
        public void Show()
        {
            if (_view.ReorderParams.Target == null) { return; }

            if (_view.ReorderParams.Parameters.Count < 2) 
            {
                var message = string.Format(RubberduckUI.ReorderPresenter_LessThanTwoParametersError, _view.ReorderParams.Target.IdentifierName);
                MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; 
            }

            _view.InitializeParameterGrid();
            _view.ShowDialog();
        }

        /// <summary>
        /// Handler for OK button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OkButtonClicked(object sender, EventArgs e)
        {
            if (!_view.ReorderParams.Parameters.Where((t, i) => t.Index != i).Any() || !IsValidParamOrder())
            {
                return;
            }

            _view.ReorderParams.Parameters = _view.ReorderParams.Parameters;
            _view.ReorderParams.Refactor();
        }

        private bool IsValidParamOrder()
        {
            var indexOfFirstOptionalParam = _view.ReorderParams.Parameters.FindIndex(param => param.IsOptional);
            if (indexOfFirstOptionalParam >= 0)
            {
                for (var index = indexOfFirstOptionalParam + 1; index < _view.ReorderParams.Parameters.Count; index++)
                {
                    if (!_view.ReorderParams.Parameters.ElementAt(index).IsOptional)
                    {
                        MessageBox.Show(RubberduckUI.ReorderPresenter_OptionalParametersMustBeLastError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }

            var indexOfParamArray = _view.ReorderParams.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray >= 0)
            {
                if (indexOfParamArray != _view.ReorderParams.Parameters.Count - 1)
                {
                    MessageBox.Show(RubberduckUI.ReorderPresenter_ParamArrayError, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Displays a prompt asking the user whether the method signature should be adjusted
        /// if the target declaration implements an interface method.
        /// </summary>
        private Declaration PromptIfTargetImplementsInterface()
        {
            var declaration = _view.ReorderParams.Target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (declaration == null || interfaceImplementation == null)
            {
                return declaration;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.ReorderPresenter_TargetIsInterfaceMemberImplementation, declaration.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                return null;
            }

            return interfaceMember;
        }
    }
}
