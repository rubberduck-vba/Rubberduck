using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ReferenceBrowserViewModel : ViewModelBase
    {
        private readonly VBE _vbe;
        private readonly RegisteredLibraryModelService _service;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _vbaProjectReferences;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _registeredComReferences;
        private string _filter;

        public ReferenceBrowserViewModel(VBE vbe, RegisteredLibraryModelService service)
        {
            _vbe = vbe;
            _service = service;
            
            _registeredComReferences = new ObservableCollection<RegisteredLibraryViewModel>();
            ComReferences = new CollectionViewSource {Source = _registeredComReferences}.View;
            //ComReferences.DeferRefresh();
            ComReferences.SortDescriptions.Add(
                new SortDescription(nameof(RegisteredLibraryViewModel.CanRemoveReference),
                ListSortDirection.Ascending));
            ComReferences.SortDescriptions.Add(
                new SortDescription(nameof(RegisteredLibraryViewModel.IsActiveProjectReference),
                ListSortDirection.Descending));
            ComReferences.SortDescriptions.Add(
                new SortDescription(nameof(RegisteredLibraryViewModel.Name), 
                ListSortDirection.Ascending));
            //ComReferences.Refresh();

            _vbaProjectReferences = new ObservableCollection<RegisteredLibraryViewModel>();
            VbaProjectReferences = new CollectionViewSource {Source = _vbaProjectReferences }.View;

            BuildTypeLibraryReferenceViewModels();
            BuildVbaProjectReferenceViewModels();
        }

        public ICollectionView ComReferences { get; }

        public ICollectionView VbaProjectReferences { get; }

        public string ComReferencesFilter
        {
            get { return _filter; }
            set
            {
                if (value == _filter)
                {
                    return;
                }
                _filter = value;
                FilterComReferences();
                OnPropertyChanged();
            }
        }

        private void FilterComReferences()
        {
            if (string.IsNullOrWhiteSpace(_filter))
            {
                ComReferences.Filter = null;
            }
            else
            {
                ComReferences.Filter = o => 
                    ((RegisteredLibraryViewModel) o).Name.ToLower()
                    .Contains(_filter.ToLower());
            }
        }

        private void BuildTypeLibraryReferenceViewModels()
        {
            var list = _service.GetAllRegisteredLibraries()
                .Select(l => new RegisteredLibraryViewModel(l, _vbe.ActiveVBProject));

            foreach (var vm in list)
            {
                _registeredComReferences.Add(vm);
            }
        }

        public void BuildVbaProjectReferenceViewModels()
        {
            var vbaReferences = _vbe.ActiveVBProject.References
                .OfType<Reference>()
                .Where(r => r.Type == vbext_RefKind.vbext_rk_Project);

            foreach (var reference in vbaReferences)
            {
                var model = new VbaProjectReferenceModel(reference);
                var vm = new RegisteredLibraryViewModel(model, _vbe.ActiveVBProject);
                _vbaProjectReferences.Add(vm);
            }
        }
    }
}
