using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Data;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Interaction;
using Rubberduck.Resources;
using Rubberduck.UI.Command;
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
        private readonly IMessageBox _messageBox;
        
        private readonly ObservableCollection<ReferenceModel> _available;
        private readonly ObservableCollection<ReferenceModel> _project;

        public AddRemoveReferencesViewModel(IAddRemoveReferencesModel model, IMessageBox messageBox)
        {
            Model = model;
            _messageBox = messageBox;

            foreach (var reference in model.References
                .Where(item => Model.Settings.PinnedReferences
                    .Contains(item.Type == ReferenceKind.TypeLibrary ? item.Guid.ToString() : item.FullPath)))
            {
                reference.IsPinned = true;
            }

            foreach (var reference in model.References
                .Where(item => Model.Settings.RecentReferences
                    .Contains(item.Type == ReferenceKind.TypeLibrary ? item.Guid.ToString() : item.FullPath)))
            {
                reference.IsRecent = true;
            }

            _available = new ObservableCollection<ReferenceModel>(model.References
                .Where(reference => !reference.IsReferenced).OrderBy(reference => reference.Description));
            _project = new ObservableCollection<ReferenceModel>(model.References
                .Where(reference => reference.IsReferenced).OrderBy(reference => reference.Priority));

            AddCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteAddCommand);
            RemoveCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveCommand);
            BrowseCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteBrowseCommand);
            MoveUpCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteMoveUpCommand);
            MoveDownCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteMoveDownCommand);
            PinLibraryCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecutePinLibraryCommand);
            PinReferenceCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecutePinReferenceCommand);
        }

        public IAddRemoveReferencesModel Model { get; set; }

        public ICommand AddCommand { get; }

        public ICommand RemoveCommand { get; }
        /// <summary>
        /// Prompts user for a .tlb, .dll, or .ocx file, and attempts to append it to <see cref="ProjectReferences"/>.
        /// </summary>
        public ICommand BrowseCommand { get; }

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
            ProjectReferences.Refresh();
        }
     
        private static readonly List<string> FileFilters = new List<string>
        {
            RubberduckUI.References_BrowseFilterExecutable,
            RubberduckUI.References_BrowseFilterExcel,
            RubberduckUI.References_BrowseFilterTypes,
            RubberduckUI.References_BrowseFilterActiveX,
            RubberduckUI.References_BrowseFilterAllFiles,
        };

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

                var project = Model.Project.Project;
                using (var references = project.References)
                {
                    try
                    {
                        using (var reference = references.AddFromFile(dialog.FileName))
                        {
                            if (reference is null)
                            {
                                return;
                            }

                            _project.Add(existing ?? new ReferenceModel(reference, _project.Count + 1));
                            ProjectReferences.Refresh();
                            if (existing is null)
                            {
                                return;
                            }

                            existing.Priority = _project.Count + 1;
                            _available.Remove(existing);
                            AvailableReferences.Refresh();
                        }
                    }
                    catch (COMException ex)
                    {
                        _messageBox.NotifyWarn(ex.Message, RubberduckUI.References_AddFailedCaption);
                    }
                }
            }
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
    }
}
