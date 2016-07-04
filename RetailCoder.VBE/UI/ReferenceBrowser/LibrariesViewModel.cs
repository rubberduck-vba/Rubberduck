using System.Collections.ObjectModel;

namespace Rubberduck.UI.ReferenceBrowser
{
    public abstract class LibrariesViewModel : ViewModelBase
    {
        protected readonly ObservableCollection<RegisteredLibraryViewModel> _registeredLibraries;
        private RegisteredLibraryViewModel _selectedLibrary;

        internal LibrariesViewModel()
        {
            _registeredLibraries = new ObservableCollection<RegisteredLibraryViewModel>();
            RegisteredLibraries = new ReadOnlyObservableCollection<RegisteredLibraryViewModel>(_registeredLibraries);
        }

        public ReadOnlyObservableCollection<RegisteredLibraryViewModel> RegisteredLibraries { get; }

        public RegisteredLibraryViewModel SelectedLibrary
        {
            get { return _selectedLibrary; }
            set
            {
                _selectedLibrary = value;
                OnPropertyChanged();
            }
        }
    }
}