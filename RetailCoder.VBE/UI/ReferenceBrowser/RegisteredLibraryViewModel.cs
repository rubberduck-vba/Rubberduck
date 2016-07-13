namespace Rubberduck.UI.ReferenceBrowser
{
    public class RegisteredLibraryViewModel : ViewModelBase
    {
        private bool _isActiveReference;
        private bool _referenceIsRemovable = true;

        public RegisteredLibraryViewModel(AbstractReferenceModel model, bool isActiveReference, bool canRemoveReference)
        {
            Model = model;
            IsActiveProjectReference = isActiveReference;
            CanRemoveReference = canRemoveReference;
        }

        public AbstractReferenceModel Model { get; private set; }

        public bool IsActiveProjectReference
        {
            get { return _isActiveReference; }
            set
            {
                if (value == _isActiveReference)
                {
                    return;
                }
                _isActiveReference = value;
                OnPropertyChanged();
            }
        }

        public bool CanRemoveReference
        {
            get { return _referenceIsRemovable; }
            set
            {
                if (value == _referenceIsRemovable)
                {
                    return;
                }
                _referenceIsRemovable = value;
                OnPropertyChanged();
            }
        }
    }
}