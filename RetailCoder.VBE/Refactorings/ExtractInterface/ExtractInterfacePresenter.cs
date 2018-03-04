using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public interface IExtractInterfacePresenter
    {
        ExtractInterfaceModel Show();
    }

    public class ExtractInterfacePresenter : IExtractInterfacePresenter
    {
        private readonly IRefactoringDialog<ExtractInterfaceViewModel> _view;
        private readonly ExtractInterfaceModel _model;

        public ExtractInterfacePresenter(IRefactoringDialog<ExtractInterfaceViewModel> view, ExtractInterfaceModel model)
        {
            _view = view;
            _model = model;
        }

        public ExtractInterfaceModel Show()
        {
            if (_model.TargetDeclaration == null)
            {
                return null;
            }

            _view.ViewModel.ComponentNames = _model.State.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .Where(moduleDeclaration => moduleDeclaration.ProjectId == _model.TargetDeclaration.ProjectId)
                .Select(module => module.ComponentName)
                .ToList();
            _view.ViewModel.InterfaceName = _model.InterfaceName;
            _view.ViewModel.Members = _model.Members.ToList();

            _view.ShowDialog();
            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }

            _model.InterfaceName = _view.ViewModel.InterfaceName;
            _model.Members = _view.ViewModel.Members;
            return _model;
        }
    }
}
