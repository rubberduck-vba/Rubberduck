using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ReferenceBrowserViewModel : ViewModelBase
    {
        private readonly VBE _vbe;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _libraryViewModels;
        private readonly ReadOnlyObservableCollection<RegisteredLibraryViewModel> _libraryViewModelsReadOnly;
        private RegisteredLibraryViewModel _selectedLibrary;

        public ReferenceBrowserViewModel(VBE vbe, RegisteredLibraryModelService service)
        {
            _vbe = vbe;

            _libraryViewModels = 
                new ObservableCollection<RegisteredLibraryViewModel>();
            _libraryViewModelsReadOnly = 
                new ReadOnlyObservableCollection<RegisteredLibraryViewModel>(_libraryViewModels);

            var allLibraries = service.GetAllRegisteredLibraries();
            BuildViewModels(allLibraries);
        }

        public ReadOnlyObservableCollection<RegisteredLibraryViewModel> RegisteredLibraries
        {
            get { return _libraryViewModelsReadOnly; }
        }

        public RegisteredLibraryViewModel SelectedLibrary
        {
            get { return _selectedLibrary; }
            set
            {
                _selectedLibrary = value;
                OnPropertyChanged();
            }
        }

        private async void BuildViewModels(IEnumerable<RegisteredLibraryModel> libraries)
        {
            var list = libraries
                .Select(l => new RegisteredLibraryViewModel(l, _vbe.ActiveVBProject))
                .ToList();

            // Sometimes the sort can take a while.  Lets do it async.
            await Task.Run(() => list.Sort());

            foreach (var vm in list)
            {
                _libraryViewModels.Add(vm);
            }
        }
    }
}
