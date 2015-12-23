namespace Rubberduck.Refactorings.ExtractInterface
{
    public interface IExtractInterfacePresenter
    {
        ExtractInterfaceModel Show();
    }

    public class ExtractInterfacePresenter : IExtractInterfacePresenter
    {
        private readonly IExtractInterfaceView _view;
        private readonly ExtractInterfaceModel _model;

        public ExtractInterfacePresenter(IExtractInterfaceView view, ExtractInterfaceModel model)
        {
            _view = view;
            _model = model;
        }

        public ExtractInterfaceModel Show()
        {
            if (_model.TargetDeclaration == null) { return null; }

            _view.InterfaceName = "I" + _model.TargetDeclaration.IdentifierName;
            _view.Members = _model.Members;

            _view.ShowDialog();
            return null;
        }
    }
}