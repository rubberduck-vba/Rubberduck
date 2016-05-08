using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public interface IExtractInterfacePresenter
    {
        ExtractInterfaceModel Show();
    }

    public class ExtractInterfacePresenter : IExtractInterfacePresenter
    {
        private readonly IExtractInterfaceDialog _view;
        private readonly ExtractInterfaceModel _model;

        public ExtractInterfacePresenter(IExtractInterfaceDialog view, ExtractInterfaceModel model)
        {
            _view = view;
            _model = model;
        }

        public ExtractInterfaceModel Show()
        {
            if (_model.TargetDeclaration == null) { return null; }

            _view.ComponentNames =
                _model.TargetDeclaration.Project.VBComponents.Cast<VBComponent>().Select(c => c.Name).ToList();
            _view.InterfaceName = _model.InterfaceName;
            _view.Members = _model.Members;

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.InterfaceName = _view.InterfaceName;
            _model.Members = _view.Members;
            return _model;
        }
    }
}