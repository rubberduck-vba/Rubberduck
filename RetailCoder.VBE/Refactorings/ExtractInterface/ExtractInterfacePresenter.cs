using System.Linq;
using System.Windows.Forms;
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

            _view.ViewModel.ComponentNames = _model.TargetDeclaration.Project.VBComponents.Select(c => c.Name).ToList();
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
