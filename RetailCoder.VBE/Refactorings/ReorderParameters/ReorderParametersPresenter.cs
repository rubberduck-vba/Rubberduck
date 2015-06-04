using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenter
    {
        private readonly IReorderParametersView _view;
        private readonly ReorderParametersModel _model;

        public ReorderParametersPresenter(IReorderParametersView view, ReorderParametersModel model)
        {
            _view = view;
            _model = model;
        }

        public ReorderParametersModel Show()
        {
            _model.TargetDeclaration = PromptIfTargetImplementsInterface();
            _model.LoadParameters();

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
            if (confirm == DialogResult.No)
            {
                return null;
            }

            return interfaceMember;
        }
    }
}
