using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ReferenceBrowserViewModel : ViewModelBase, IDisposable
    {
        private readonly VBE _vbe;
        private readonly RegisteredLibraryModelService _service;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _vbaProjectReferences;
        private readonly ObservableCollection<RegisteredLibraryViewModel> _registeredComReferences;
        private string _filter;

        public ReferenceBrowserViewModel(VBE vbe, RegisteredLibraryModelService service, IOpenFileDialog filePicker)
        {
            _vbe = vbe;
            _service = service;

            _registeredComReferences = new ObservableCollection<RegisteredLibraryViewModel>();
            ComReferences = new CollectionViewSource {Source = _registeredComReferences}.View;
            //ComReferences.DeferRefresh();  would prefer to use this for performance but gives an error.
            ComReferences.SortDescriptions.Add(new SortDescription("CanRemoveReference", ListSortDirection.Ascending));
            ComReferences.SortDescriptions.Add(new SortDescription("IsActiveProjectReference", ListSortDirection.Descending));
            ComReferences.SortDescriptions.Add(new SortDescription("Name", ListSortDirection.Ascending));
            //ComReferences.Refresh();

            _vbaProjectReferences = new ObservableCollection<RegisteredLibraryViewModel>();
            VbaProjectReferences = new CollectionViewSource {Source = _vbaProjectReferences }.View;

            BuildTypeLibraryReferenceViewModels();
            BuildVbaProjectReferenceViewModels();

            AddVbaProjectReferenceCommand = new AddReferenceCommand(filePicker, AddVbaReference);
        }

        public ICollectionView ComReferences { get; private set; }

        public ICollectionView VbaProjectReferences { get; private set; }

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

        public ICommand AddVbaProjectReferenceCommand { get; private set; }

        private void AddVbaReference(string filePath)
        {
            // TODO this may throw exceptions.  Gotta to catch 'em all!
            var reference = _vbe.ActiveVBProject.References.AddFromFile(filePath);
            CreateViewModelForVbaProjectReference(reference);
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

        private void BuildVbaProjectReferenceViewModels()
        {
            var vbaReferences = _vbe.ActiveVBProject.References
                .OfType<Reference>()
                .Where(r => r.Type == vbext_RefKind.vbext_rk_Project);

            foreach (var reference in vbaReferences)
            {
                CreateViewModelForVbaProjectReference(reference);
            }
        }

        private void CreateViewModelForVbaProjectReference(Reference reference)
        {
            var model = new VbaProjectReferenceModel(reference);
            var vm = new RegisteredLibraryViewModel(model, _vbe.ActiveVBProject);
            _vbaProjectReferences.Add(vm);
        }

        public void Dispose()
        {

            var command = AddVbaProjectReferenceCommand as IDisposable;
            if (command != null)
            {
                command.Dispose();
            }
            AddVbaProjectReferenceCommand = null;
        }

        private class AddReferenceCommand : CommandBase, IDisposable
        {
            private readonly IOpenFileDialog _filePicker;
            private Action<string> _addReferenceCallback;

            internal AddReferenceCommand(IOpenFileDialog filePicker, Action<string> addReferenceCallback) 
                : base(LogManager.GetCurrentClassLogger())
            {
                if (addReferenceCallback == null)
                {
                    throw new ArgumentNullException("addReferenceCallback");
                }
                _addReferenceCallback = addReferenceCallback;
                _filePicker = filePicker;
                _filePicker.AddExtension = true;
                _filePicker.AutoUpgradeEnabled = true;
                _filePicker.CheckFileExists = true;
                _filePicker.Multiselect = false;
                _filePicker.ShowHelp = false;
                _filePicker.Filter = @"Excel Files|*.xls;*.xlsx;*.xlsm";
                _filePicker.CheckFileExists = true;
            }

            protected override void ExecuteImpl(object parameter)
            {
                if (_filePicker.ShowDialog() == DialogResult.OK)
                {
                    _addReferenceCallback(_filePicker.FileNames[0]);
                }
            }

            public void Dispose()
            {
                if (_filePicker != null)
                {
                    _filePicker.Dispose();
                }
                _addReferenceCallback = null;
            }
        }
    }
}
