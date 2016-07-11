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
        private readonly ObservableCollection<RegisteredLibraryViewModel> _projectReferences;
        private string _filter;

        public ReferenceBrowserViewModel(VBE vbe, RegisteredLibraryModelService service, IOpenFileDialog filePicker)
        {
            _vbe = vbe;
            _service = service;

            _projectReferences = new ObservableCollection<RegisteredLibraryViewModel>();
            ComReferences = new CollectionViewSource {Source = _projectReferences }.View;
            FilterComReferences();

            VbaProjectReferences = new CollectionViewSource {Source = _projectReferences }.View;
            VbaProjectReferences.Filter = o => ((RegisteredLibraryViewModel) o).Model is VbaProjectReferenceModel;

            BuildViewModels();

            AddVbaProjectReferenceCommand = null;// new AddReferenceCommand(filePicker, AddVbaReference);
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => OnCloseWindow(new EventArgs()));
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdateReferencesAndClose());
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

        public ICommand CancelButtonCommand { get; private set; }

        public ICommand OkButtonCommand { get; private set; }

        private void BuildViewModels()
        {
            var libraries = _service.GetAllRegisteredLibraries();
            var currentReferences = _vbe.ActiveVBProject.References.OfType<Reference>().ToList();

            // first add the existing references to the list.  That way we have the correct ordering.
            foreach (var r in currentReferences)
            {
                // create a view model for the given reference.
                if (r.Type == vbext_RefKind.vbext_rk_Project)
                {
                    var model = new VbaProjectReferenceModel(r);
                    var vm = new RegisteredLibraryViewModel(model, true, !r.BuiltIn);
                    _projectReferences.Add(vm);
                }
                else
                {
                    var model = libraries.FirstOrDefault(l => string.Equals(l.FilePath, r.FullPath, StringComparison.CurrentCultureIgnoreCase));
                    libraries.Remove(model);  // remove here so we can sort and add later.
                    var vm = new RegisteredLibraryViewModel(model, true, !r.BuiltIn);
                    _projectReferences.Add(vm);
                }
            }

            // add the remaining registered libraries to the collection in sorted order.
            libraries.Sort((f, s) => string.Compare(f.Name, s.Name, StringComparison.CurrentCultureIgnoreCase));
            foreach (var l in libraries)
            {
                var vm = new RegisteredLibraryViewModel(l, false, true);
                _projectReferences.Add(vm);
            }
        }

        private void UpdateReferencesAndClose()
        {
            UpdateReferences();
            OnCloseWindow(new EventArgs());
        }

        private void UpdateReferences()
        {
            var references = _vbe.ActiveVBProject.References;
            var activeModels = _projectReferences
                .Where(r => r.IsActiveProjectReference)
                .Select(r => r.Model)
                .ToList();

            // remove any references which aren't in the active models list.
            foreach (var r in references.OfType<Reference>().ToList())
            {
                if (!activeModels.Any(m => string.Equals(m.FilePath, r.FullPath, StringComparison.CurrentCultureIgnoreCase)))
                {
                    references.Remove(r);
                }
            }

            var referenceIndex = 1;
            var modelIndex = 0;
            while (referenceIndex <= references.Count)
            {
                var currentModel = activeModels[modelIndex];
                var currentReference = references.Item(referenceIndex);

                if (string.Equals(currentModel.FilePath, currentReference.FullPath, StringComparison.CurrentCultureIgnoreCase))
                {
                    modelIndex++;
                    referenceIndex++;
                }
                else
                {
                    references.Remove(currentReference);
                }
            }
            // add the remaining models to the references.
            while (modelIndex < activeModels.Count)
            {
                var currentModel = activeModels[modelIndex];
                references.AddFromFile(currentModel.FilePath);
                modelIndex++;
            }
        }

        private void FilterComReferences()
        {
            if (string.IsNullOrWhiteSpace(_filter))
            {
                ComReferences.Filter = o => ((RegisteredLibraryViewModel) o).Model is RegisteredLibraryModel;
            }
            else
            {
                ComReferences.Filter = o =>
                {
                    var model = ((RegisteredLibraryViewModel)o).Model;
                    return model.Name.ToLower().Contains(_filter.ToLower())
                        && model is RegisteredLibraryModel;
                };
            }
        }

        public event EventHandler<EventArgs> CloseWindow; 

        private void OnCloseWindow(EventArgs args)
        {
            var handler = CloseWindow;
            if (handler != null)
            {
                handler.Invoke(this, args);
            }
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
