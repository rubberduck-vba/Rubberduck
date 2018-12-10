using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.ComponentModel;
using System.IO;
using System.Windows.Data;
using System.Windows.Forms;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Resources;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.AddRemoveReferences
{
    public enum ReferenceFilter
    {
        Recent,
        Pinned,
        ComTypes,
        Projects
    }

    public class AddRemoveReferencesViewModel : ViewModelBase
    {
        private static readonly Dictionary<string, string[]> HostFileFilters = new Dictionary<string, string[]>
        {
            { "EXCEL.EXE", new [] {"xlsm","xlam","xls","xla"} },
            { "WINWORD.EXE", new [] {"docm","dotm","doc","dot"} },
            { "MSACCESS.EXE", new [] {"accdb","mdb","mda","mde","accde"} },
            { "POWERPNT.EXE", new [] {"ppam","ppa"} },
            // TODO
            //{ "OUTLOOK.EXE", new [] {"?"} },
            //{ "WINPROJ.EXE",  new [] {"?"} },
            //{ "MSPUB.EXE",  new [] {"?"} },
            //{ "VISIO.EXE",  new [] {"?"} },
            //{ "ACAD.EXE",  new [] {"?"} },
            //{ "CORELDRW.EXE",  new [] {"?"} }
        };

        private static readonly Dictionary<string, string> HostFilters = new Dictionary<string, string>
        {
            { "EXCEL.EXE", string.Format(RubberduckUI.References_BrowseFilterExcel, string.Join(";", HostFileFilters["EXCEL.EXE"].Select(_ => $"*.{_}"))) },
            { "WINWORD.EXE", string.Format(RubberduckUI.References_BrowseFilterWord, string.Join(";", HostFileFilters["WINWORD.EXE"].Select(_ => $"*.{_}"))) }, 
            { "MSACCESS.EXE", string.Format(RubberduckUI.References_BrowseFilterAccess, string.Join(";", HostFileFilters["MSACCESS.EXE"].Select(_ => $"*.{_}"))) },
            { "POWERPNT.EXE", string.Format(RubberduckUI.References_BrowseFilterPowerPoint, string.Join(";", HostFileFilters["POWERPNT.EXE"].Select(_ => $"*.{_}"))) },
        };

        private static readonly List<string> FileFilters = new List<string>
        {
            RubberduckUI.References_BrowseFilterExecutable,
            RubberduckUI.References_BrowseFilterTypes,
            RubberduckUI.References_BrowseFilterActiveX,
            RubberduckUI.References_BrowseFilterAllFiles,
        };

        private static bool HostHasProjects { get; }

        static AddRemoveReferencesViewModel()
        {
            var host = Path.GetFileName(Application.ExecutablePath).ToUpperInvariant();
            if (!HostFilters.ContainsKey(host))
            {
                return;
            }

            HostHasProjects = true;
            FileFilters.Insert(0, HostFilters[host]);
        }

        public event EventHandler<DialogResult> OnWindowClosed;
        private void CloseWindowOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);
        private void CloseWindowCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);

        private readonly List<(int?, ReferenceInfo)> _clean;
        private readonly ObservableCollection<ReferenceModel> _available;
        private readonly ObservableCollection<ReferenceModel> _project;
        private readonly IReferenceReconciler _reconciler;

        public AddRemoveReferencesViewModel(IAddRemoveReferencesModel model, IReferenceReconciler reconciler)
        {
            Model = model;
            _reconciler = reconciler;

            _available = new ObservableCollection<ReferenceModel>(model.References
                .Where(reference => !reference.IsReferenced).OrderBy(reference => reference.Description));
            _project = new ObservableCollection<ReferenceModel>(model.References
                .Where(reference => reference.IsReferenced).OrderBy(reference => reference.Priority));

            _clean = new List<(int?, ReferenceInfo)>(_project.Select(reference => (reference.Priority, reference.ToReferenceInfo())));

            BuiltInReferenceCount = _project.Count(reference => reference.IsBuiltIn);

            AddCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteAddCommand);
            RemoveCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveCommand);
            BrowseCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteBrowseCommand);
            MoveUpCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteMoveUpCommand);
            MoveDownCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteMoveDownCommand);
            PinLibraryCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecutePinLibraryCommand);
            PinReferenceCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecutePinReferenceCommand);
           
            OkCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseWindowOk());
            CancelCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseWindowCancel());
            ApplyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteApplyCommand);
        }

        public IAddRemoveReferencesModel Model { get; set; }

        public bool ProjectsVisible => HostHasProjects;

        public int BuiltInReferenceCount { get; }

        public ICommand AddCommand { get; }

        public ICommand RemoveCommand { get; }
        /// <summary>
        /// Prompts user for a .tlb, .dll, or .ocx file, and attempts to append it to <see cref="ProjectReferences"/>.
        /// </summary>
        public ICommand BrowseCommand { get; }

        public CommandBase OkCommand { get; }
        public CommandBase CancelCommand { get; }

        /// <summary>
        /// Applies all changes to project references.
        /// </summary>
        public ICommand ApplyCommand { get; }

        /// <summary>
        /// Moves the <see cref="SelectedReference"/> up on the 'Priority' tab.
        /// </summary>
        public ICommand MoveUpCommand { get; }

        /// <summary>
        /// Moves the <see cref="SelectedReference"/> down on the 'Priority' tab.
        /// </summary>
        public ICommand MoveDownCommand { get; }

        public ICommand PinLibraryCommand { get; }

        public ICommand PinReferenceCommand { get; }

        private void ExecuteAddCommand(object parameter)
        {
            if (SelectedLibrary == null)
            {
                return;
            }

            SelectedLibrary.Priority = _project.Count + 1;
            _project.Add(SelectedLibrary);
            EvaluateProjectDirty();
            ProjectReferences.Refresh();
            _available.Remove(SelectedLibrary);
        }

        private void ExecuteRemoveCommand(object parameter)
        {
            if (SelectedReference == null)
            {
                return;
            }

            var priority = SelectedReference.Priority;
            SelectedReference.Priority = null;
            _available.Add(SelectedReference);            
            _project.Remove(SelectedReference);

            foreach (var reference in _project.Where(lib => lib.Priority > priority).ToList())
            {
                reference.Priority--;
            }

            EvaluateProjectDirty();
            ProjectReferences.Refresh();
        }
        
        private void ExecuteBrowseCommand(object parameter)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = string.Join("|", FileFilters),
                Title = RubberduckUI.References_BrowseCaption
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName))
                {
                    return;
                }

                var existing = _available.FirstOrDefault(library =>
                    library.FullPath.Equals(dialog.FileName, StringComparison.OrdinalIgnoreCase));

                if (existing is null)
                {
                    var adding = _reconciler.GetLibraryInfoFromPath(dialog.FileName);
                    if (adding is null)
                    {
                        return;
                    }

                    adding.Priority = _project.Count + 1;
                    Model.References.Add(adding);
                    _project.Add(adding);
                }
                else
                {
                    _project.Add(existing);
                    existing.Priority = _project.Count + 1;
                    _available.Remove(existing);
                    AvailableReferences.Refresh();
                }
                ProjectReferences.Refresh();
            }
        }

        private void ExecuteApplyCommand(object parameter)
        {
            var changed = _reconciler.ReconcileReferences(Model, _available.ToList());
            foreach (var reference in changed.Where(reference => !_project.Contains(reference)).ToList())
            {
                _project.Add(reference);
            }
            
            AvailableReferences.Refresh();
            ProjectReferences.Refresh();
        }

        private void ExecuteMoveUpCommand(object parameter)
        {
            if (SelectedReference == null || SelectedReference.IsBuiltIn || SelectedReference.Priority == 1)
            {
                return;
            }

            var swap = _project.SingleOrDefault(reference => reference.Priority == SelectedReference.Priority - 1);

            if (swap is null || swap.IsBuiltIn)
            {
                return;
            }

            swap.Priority = SelectedReference.Priority;
            SelectedReference.Priority--;
            EvaluateProjectDirty();
            ProjectReferences.Refresh();
        }

        private void ExecuteMoveDownCommand(object parameter)
        {
            if (SelectedReference == null || SelectedReference.IsBuiltIn || SelectedReference.Priority == _project.Count)
            {
                return;
            }

            var swap = _project.SingleOrDefault(reference => reference.Priority == SelectedReference.Priority + 1);

            if (swap is null || swap.IsBuiltIn)
            {
                return;
            }

            swap.Priority = SelectedReference.Priority;
            SelectedReference.Priority++;
            EvaluateProjectDirty();
            ProjectReferences.Refresh();
        }

        private void ExecutePinLibraryCommand(object parameter)
        {
            if (SelectedLibrary == null)
            {
                return;
            }
            SelectedLibrary.IsPinned = !SelectedLibrary.IsPinned;
            AvailableReferences.Refresh();
        }

        private void ExecutePinReferenceCommand(object parameter)
        {
            if (SelectedReference == null || SelectedReference.IsBuiltIn)
            {
                return;
            }
            SelectedReference.IsPinned = !SelectedReference.IsPinned;
            ProjectReferences.Refresh();
        }

        public ICollectionView ProjectReferences
        {
            get
            {
                var projects = CollectionViewSource.GetDefaultView(_project);
                projects.SortDescriptions.Add(new SortDescription("Priority", ListSortDirection.Ascending));
                return projects;
            }
        }

        public ICollectionView AvailableReferences
        {
            get
            {
                var available = CollectionViewSource.GetDefaultView(_available);
                available.Filter = reference => Filter((ReferenceModel)reference);
                return available;
            }
        }

        private string _filter;
        public string SelectedFilter
        {
            get => _filter;
            set
            {
                _filter = value;
                AvailableReferences.Refresh();
            }
        }

        private bool Filter(ReferenceModel reference)
        {
            var filtered = false;
            Enum.TryParse<ReferenceFilter>(SelectedFilter, out var filter);
            switch (filter)
            {
                case ReferenceFilter.Recent:
                    filtered = reference.IsRecent;
                    break;
                case ReferenceFilter.Pinned:
                    filtered = reference.IsPinned;
                    break;
                case ReferenceFilter.ComTypes:
                    filtered = reference.Type == ReferenceKind.TypeLibrary;
                    break;
                case ReferenceFilter.Projects:
                    filtered = reference.Type == ReferenceKind.Project;
                    break;
            }

            var searched = string.IsNullOrEmpty(Search)
                           || reference.Name.IndexOf(Search, StringComparison.OrdinalIgnoreCase) >= 0
                           || reference.Description.IndexOf(Search, StringComparison.OrdinalIgnoreCase) >= 0
                           || reference.FullPath.IndexOf(Search, StringComparison.OrdinalIgnoreCase) >= 0;

            return filtered && searched;
        }

        private string _search = string.Empty;
        public string Search
        {
            get => _search;
            set
            {
                _search = value;
                AvailableReferences.Refresh();
            }
        }

        private ReferenceModel _selection;
        public ReferenceModel CurrentSelection
        {
            get => _selection;
            set
            {
                _selection = value;
                OnPropertyChanged();
            }
        }

        private ReferenceModel _reference;
        public ReferenceModel SelectedReference
        {
            get => _reference;
            set
            {
                _reference = value;
                CurrentSelection = _reference;
            }
        }

        private ReferenceModel _library;
        public ReferenceModel SelectedLibrary
        {
            get => _library;
            set
            {
                _library = value;
                CurrentSelection = _library;
            }
        }

        private bool _dirty;
        public bool IsProjectDirty
        {
            get => _dirty;
            set
            {
                _dirty = value;
                OnPropertyChanged();
            }
        }

        private void EvaluateProjectDirty()
        {
            var selected = _project.Select(reference => (reference.Priority, reference.ToReferenceInfo())).ToList();
            IsProjectDirty = selected.Count != _clean.Count || !_clean.All(selected.Contains);
        }
    }
}
