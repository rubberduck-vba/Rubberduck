using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Rubberduck.Navigation.CodeExplorer;
using System.Windows;
using Rubberduck.Navigation.Folders;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public class CodeMetricsViewModel : ViewModelBase, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly ICodeMetricsAnalyst _analyst;
        private readonly FolderHelper _folderHelper;
        private readonly IVBE _vbe;

        public CodeMetricsViewModel(RubberduckParserState state, ICodeMetricsAnalyst analyst, FolderHelper folderHelper, IVBE vbe)
        {
            _state = state;
            _analyst = analyst;
            _folderHelper = folderHelper;
            _state.StateChanged += OnStateChanged;
            _vbe = vbe;
        }
        
        private void OnStateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready && e.State != ParserState.Error && e.State != ParserState.ResolverError && e.State != ParserState.UnexpectedError)
            {
                IsBusy = true;
            }

            if (e.State == ParserState.Ready)
            {
                UpdateData();
                IsBusy = false;
            }

            if (e.State == ParserState.Error || e.State == ParserState.ResolverError || e.State == ParserState.UnexpectedError)
            {
                IsBusy = false;
            }
        }

        private void UpdateData()
        {
            IsBusy = true;

            var metricResults = _analyst.GetMetrics(_state);
            resultsByDeclaration = metricResults.GroupBy(r => r.Declaration).ToDictionary(g => g.Key, g => g.ToList());

            if (Projects == null)
            {
                Projects = new ObservableCollection<CodeExplorerItemViewModel>();
            }

            IsBusy = _state.Status != ParserState.Pending && _state.Status <= ParserState.ResolvedDeclarations;

            var userDeclarations = _state.DeclarationFinder.AllUserDeclarations
                .GroupBy(declaration => declaration.ProjectId)
                .ToList();

            var newProjects = userDeclarations
                .Where(grouping => grouping.Any(declaration => declaration.DeclarationType == DeclarationType.Project))
                .Select(grouping =>
                new CodeExplorerProjectViewModel(_folderHelper,
                    grouping.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Project),
                    grouping,
                    _vbe)).ToList();

            UpdateNodes(Projects, newProjects);

            Projects = new ObservableCollection<CodeExplorerItemViewModel>(newProjects);

            IsBusy = false;
        }

        private void UpdateNodes(IEnumerable<CodeExplorerItemViewModel> oldList, IEnumerable<CodeExplorerItemViewModel> newList)
        {
            foreach (var item in newList)
            {
                CodeExplorerItemViewModel oldItem;

                if (item is CodeExplorerCustomFolderViewModel)
                {
                    oldItem = oldList.FirstOrDefault(i => i.Name == item.Name);
                }
                else
                {
                    oldItem = oldList.FirstOrDefault(i =>
                        item.QualifiedSelection != null && i.QualifiedSelection != null &&
                        i.QualifiedSelection.Value.QualifiedName.ProjectId ==
                        item.QualifiedSelection.Value.QualifiedName.ProjectId &&
                        i.QualifiedSelection.Value.QualifiedName.ComponentName ==
                        item.QualifiedSelection.Value.QualifiedName.ComponentName &&
                        i.QualifiedSelection.Value.Selection == item.QualifiedSelection.Value.Selection);
                }

                if (oldItem != null)
                {
                    item.IsExpanded = oldItem.IsExpanded;
                    item.IsSelected = oldItem.IsSelected;

                    if (oldItem.Items.Any() && item.Items.Any())
                    {
                        UpdateNodes(oldItem.Items, item.Items);
                    }
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }
            _isDisposed = true;

            _state.StateChanged -= OnStateChanged;
        }

        private Dictionary<Declaration, List<ICodeMetricResult>> resultsByDeclaration;

        private CodeExplorerItemViewModel _selectedItem;
        public CodeExplorerItemViewModel SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (_selectedItem == value)
                {
                    return;
                }
                _selectedItem = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(Metrics));
            }
        }

        private ObservableCollection<CodeExplorerItemViewModel> _projects;
        public ObservableCollection<CodeExplorerItemViewModel> Projects
        {
            get => _projects;
            set
            {
                _projects = new ObservableCollection<CodeExplorerItemViewModel>(value.OrderBy(o => o.NameWithSignature));

                OnPropertyChanged();
                OnPropertyChanged(nameof(TreeViewVisibility));
            }
        }

        public Visibility TreeViewVisibility => Projects == null || Projects.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
        
        public ObservableCollection<ICodeMetricResult> Metrics
        {
            get
            {
                var results = resultsByDeclaration?.FirstOrDefault(f => f.Key == SelectedItem.GetSelectedDeclaration());
                return !results.HasValue || results.Value.Value == null ? new ObservableCollection<ICodeMetricResult>() : new ObservableCollection<ICodeMetricResult>(results.Value.Value);
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                _isBusy = value;
                EmptyUIRefreshMessageVisibility = false;
                OnPropertyChanged();
            }
        }

        private bool _emptyUIRefreshMessageVisibility = true;
        public bool EmptyUIRefreshMessageVisibility
        {
            get => _emptyUIRefreshMessageVisibility;
            set
            {
                if (_emptyUIRefreshMessageVisibility != value)
                {
                    _emptyUIRefreshMessageVisibility = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
