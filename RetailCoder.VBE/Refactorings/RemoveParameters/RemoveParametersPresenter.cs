using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenter
    {
        private readonly IRemoveParametersView _view;
        private readonly RemoveParametersModel _model;

        public RemoveParametersPresenter(IRemoveParametersView view, RemoveParametersModel model)
        {
            _view = view;
            _model = model;
        }

        public RemoveParametersModel Show()
        {
            _model.TargetDeclaration = PromptIfTargetImplementsInterface();
            _model.LoadParameters();

            if (_model.Parameters.Count == 0)
            {
                var message = string.Format(RubberduckUI.RemovePresenter_NoParametersError, _model.TargetDeclaration.IdentifierName);
                MessageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }

            _view.Parameters = _model.Parameters;
            _view.InitializeParameterGrid();

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.Parameters = _view.Parameters;
            return _model;
        }

        private Declaration PromptIfTargetImplementsInterface()
        {
            var declaration = _model.TargetDeclaration;
            var interfaceImplementation = _model.Declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (declaration == null || interfaceImplementation == null)
            {
                return declaration;
            }

            var interfaceMember = _model.Declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.Refactoring_TargetIsInterfaceMemberImplementation, declaration.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            return confirm == DialogResult.No ? null : interfaceMember;
        }
    }
}
