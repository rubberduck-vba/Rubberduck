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
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
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
        public static readonly Dictionary<string, string[]> HostFileFilters = new Dictionary<string, string[]>
        {
            { "EXCEL.EXE", new [] {"xlsm","xlam","xls","xla"} },
            { "WINWORD.EXE", new [] {"docm","dotm","doc","dot"} },
            { "MSACCESS.EXE", new [] {"accdb","mdb","mda","mde","accde"} },
            { "POWERPNT.EXE", new [] {"ppam","ppa"} },
            { "OUTLOOK.EXE", new [] {"otm"} },
            { "MSPUB.EXE",  new [] {"pub"} },
            { "VISIO.EXE",  new [] {"vsd","vdx","vss","vst","vtx","vsw","vdw","vsdx","vsdm"} },
            // TODO           
            //{ "WINPROJ.EXE",  new [] {"?"} },
            //{ "ACAD.EXE",  new [] {"?"} },
            //{ "CORELDRW.EXE",  new [] {"?"} }
        };

        private static readonly Dictionary<string, string> HostFilters = new Dictionary<string, string>
        {
            { "EXCEL.EXE", string.Format(RubberduckUI.References_BrowseFilterExcel, string.Join(";", HostFileFilters["EXCEL.EXE"].Select(_ => $"*.{_}"))) },
            { "WINWORD.EXE", string.Format(RubberduckUI.References_BrowseFilterWord, string.Join(";", HostFileFilters["WINWORD.EXE"].Select(_ => $"*.{_}"))) }, 
            { "MSACCESS.EXE", string.Format(RubberduckUI.References_BrowseFilterAccess, string.Join(";", HostFileFilters["MSACCESS.EXE"].Select(_ => $"*.{_}"))) },
            { "POWERPNT.EXE", string.Format(RubberduckUI.References_BrowseFilterPowerPoint, string.Join(";", HostFileFilters["POWERPNT.EXE"].Select(_ => $"*.{_}"))) },
            { "OUTLOOK.EXE", string.Format(RubberduckUI.References_BrowseFilterOutlook, string.Join(";", HostFileFilters["OUTLOOK.EXE"].Select(_ => $"*.{_}"))) },
            { "MSPUB.EXE", string.Format(RubberduckUI.References_BrowseFilterOutlook, string.Join(";", HostFileFilters["MSPUB.EXE"].Select(_ => $"*.{_}"))) },
            { "VISIO.EXE", string.Format(RubberduckUI.References_BrowseFilterVisio, string.Join(";", HostFileFilters["VISIO.EXE"].Select(_ => $"*.{_}"))) },
        };

        private static readonly List<string> FileFilters = new List<string>
        {
            RubberduckUI.References_BrowseFilterExecutable,
            RubberduckUI.References_BrowseFilterTypes,
            RubberduckUI.References_BrowseFilterActiveX,
            RubberduckUI.References_BrowseFilterAllFiles,
        };

        public static bool HostHasProjects { get; }

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
        private void CloseWindowOk() => OnWindowClosed?.Invoke(this, IsProjectDirty ? DialogResult.OK : DialogResult.Cancel);
        private void CloseWindowCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);

        private readonly List<(int?, ReferenceInfo)> _clean;
        private readonly ObservableCollection<ReferenceModel> _available;
        private readonly ObservableCollection<ReferenceModel> _project;
        private readonly IReferenceReconciler _reconciler;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IFileSystemBrowserFactory _browser;

        public AddRemoveReferencesViewModel(
            IAddRemoveReferencesModel model, 
            IReferenceReconciler reconciler, 
            IFileSystemBrowserFactory browser,
            IProjectsProvider projectsProvider)
        {
            Model = model;
            _reconciler = reconciler;
            _browser = browser;
            _projectsProvider = projectsProvider;

            _available = new ObservableCollection<ReferenceModel>(model.References
                .Where(reference => !reference.IsReferenced).OrderBy(reference => reference.Description));
            _project = new ObservableCollection<ReferenceModel>(model.References
                .Where(reference => reference.IsReferenced).OrderBy(reference => reference.Priority));

            _clean = new List<(int?, ReferenceInfo)>(_project.Select(reference => (reference.Priority, reference.ToReferenceInfo())));

            BuiltInReferenceCount = _project.Count(reference => reference.IsBuiltIn);

            AddCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteAddCommand);
            RemoveCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveCommand);
            ClearSearchCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteClearSearchCommand);
            BrowseCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteBrowseCommand);
            MoveUpCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteMoveUpCommand);
            MoveDownCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteMoveDownCommand);
            PinLibraryCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecutePinLibraryCommand);
            PinReferenceCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecutePinReferenceCommand);
           
            OkCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseWindowOk());
            CancelCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseWindowCancel());
            ApplyCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteApplyCommand, ApplyCanExecute);
        }

        public string ProjectCaption
        {
            get
            {
                if (string.IsNullOrEmpty(Model?.Project?.IdentifierName))
                {
                    return RubberduckUI.References_Caption;
                }

                var project = _projectsProvider.Project(Model.Project.ProjectId);

                if (project == null)
                {
                    return RubberduckUI.References_Caption;
                }

                return project.ProjectDisplayName;
            }
        }

        /// <summary>
        /// The IAddRemoveReferencesModel for the view.
        /// </summary>
        public IAddRemoveReferencesModel Model { get; set; }

        /// <summary>
        /// Hides the projects filter if the host does not support them. Statically set.
        /// </summary>
        public bool ProjectsHidden => !HostHasProjects;

        /// <summary>
        /// The number of built-in (locked) references of the project.
        /// </summary>
        public int BuiltInReferenceCount { get; }

        /// <summary>
        /// Adds a reference to the project.
        /// </summary>
        public ICommand AddCommand { get; }

        /// <summary>
        /// Removes a reference from the project and makes it "available".
        /// </summary>
        public ICommand RemoveCommand { get; }

        /// <summary>
        /// Clears the search textbox.
        /// </summary>
        public ICommand ClearSearchCommand { get; }

        /// <summary>
        /// Prompts the user to browse for a reference.
        /// </summary>
        public ICommand BrowseCommand { get; }

        /// <summary>
        /// Closes the dialog and indicates changes are to be saved.
        /// </summary>
        public CommandBase OkCommand { get; }

        /// <summary>
        /// Closes the dialog and indicates changes are not to be saved.
        /// </summary>
        public CommandBase CancelCommand { get; }

        /// <summary>
        /// Applies any changes without closing the dialog.
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

        /// <summary>
        /// Pins the selected reference from the available list.
        /// </summary>
        public ICommand PinLibraryCommand { get; }

        /// <summary>
        /// Pins the selected reference from the referenced list.
        /// </summary>
        public ICommand PinReferenceCommand { get; }

        /// <summary>
        /// Delegate for AddCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
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
            AvailableReferences.Refresh();
        }

        /// <summary>
        /// Delegate for RemoveCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
        private void ExecuteRemoveCommand(object parameter)
        {
            if (SelectedReference == null || SelectedReference.IsBuiltIn)
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
            AvailableReferences.Refresh();
        }

        /// <summary>
        /// Delegate for ClearSearchCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
        private void ExecuteClearSearchCommand(object parameter)
        {
            if (!string.IsNullOrEmpty(Search))
            {
                Search = string.Empty;
            }
        }

        /// <summary>
        /// Delegate for BrowseCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
        private void ExecuteBrowseCommand(object parameter)
        {
            using (var dialog = _browser.CreateOpenFileDialog())
            {
                dialog.Filter = string.Join("|", FileFilters);
                dialog.Title = RubberduckUI.References_BrowseCaption;
                var result = dialog.ShowDialog();
                if (result != DialogResult.OK || string.IsNullOrEmpty(dialog.FileName))
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

        /// <summary>
        /// Delegate for ApplyCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
        private void ExecuteApplyCommand(object parameter)
        {
            var changed = _reconciler.ReconcileReferences(Model, _project.ToList());
            foreach (var reference in changed.Where(reference => !_project.Contains(reference)).ToList())
            {
                _project.Add(reference);
            }
            
            Model.State.OnParseRequested(this);

            _clean.Clear();
            _clean.AddRange(new List<(int?, ReferenceInfo)>(_project.Select(reference => (reference.Priority, reference.ToReferenceInfo()))));
            IsProjectDirty = false;

            AvailableReferences.Refresh();
            ProjectReferences.Refresh();
        }

        private bool ApplyCanExecute(object parameter)
        {
            return Model.State.Status == ParserState.Ready;
        }

        /// <summary>
        /// Delegate for MoveUpCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
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

        /// <summary>
        /// Delegate for MoveDownCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
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

        /// <summary>
        /// Delegate for PinLibraryCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
        private void ExecutePinLibraryCommand(object parameter)
        {
            if (SelectedLibrary == null)
            {
                return;
            }
            SelectedLibrary.IsPinned = !SelectedLibrary.IsPinned;
            AvailableReferences.Refresh();
        }

        /// <summary>
        /// Delegate for PinReferenceCommand.
        /// </summary>
        /// <param name="parameter">Ignored</param>
        private void ExecutePinReferenceCommand(object parameter)
        {
            if (SelectedReference == null)
            {
                return;
            }
            SelectedReference.IsPinned = !SelectedReference.IsPinned;
            ProjectReferences.Refresh();
        }

        /// <summary>
        /// Ordered collection of the project's currently selected references.
        /// </summary>
        public ICollectionView ProjectReferences
        {
            get
            {
                var projects = CollectionViewSource.GetDefaultView(_project);
                projects.SortDescriptions.Add(new SortDescription("Priority", ListSortDirection.Ascending));
                return projects;
            }
        }

        /// <summary>
        /// Collection of references not currently selected for the project, filtered by the current filter.
        /// </summary>
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

        /// <summary>
        /// The currently selected filter. Should be a member of ReferenceFilter.
        /// </summary>
        public string SelectedFilter
        {
            get => _filter;
            set
            {
                if (!Enum.TryParse<ReferenceFilter>(value, out _))
                {
                    return;
                } 
                _filter = value;
                AvailableReferences.Refresh();
            }
        }

        /// <summary>
        /// Applies selected filter and any search term to CollectionViewSource.
        /// </summary>
        /// <param name="reference">The ReferenceModel to test.</param>
        /// <returns>Returns true if the passed reference is included in the filtered result.</returns>
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

        /// <summary>
        /// Search term for filtering AvailableReferences.
        /// </summary>
        public string Search
        {
            get => _search;
            set
            {
                _search = value;
                OnPropertyChanged();
                AvailableReferences.Refresh();
            }
        }

        private ReferenceModel _selection;

        /// <summary>
        /// The currently selected Reference in the focused list.
        /// </summary>
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

        /// <summary>
        /// The currently selected Reference for the project.
        /// </summary>
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

        /// <summary>
        /// The currently selected available (not included in the project) Reference.
        /// </summary>
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

        /// <summary>
        /// Indicated whether any changes were made to the project's references.
        /// </summary>
        public bool IsProjectDirty
        {
            get => _dirty;
            set
            {
                _dirty = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Tests to see if any changes have been made to the project and sets IsProjectDirty to the appropriate value.
        /// </summary>
        private void EvaluateProjectDirty()
        {
            var selected = _project.Select(reference => (reference.Priority, reference.ToReferenceInfo())).ToList();
            IsProjectDirty = selected.Count != _clean.Count || !_clean.All(selected.Contains);
        }
    }
}
