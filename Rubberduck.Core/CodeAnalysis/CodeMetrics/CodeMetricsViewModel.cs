using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public sealed class CodeMetricsViewModel : ViewModelBase, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly ICodeMetricsAnalyst _analyst;
        private readonly IVBE _vbe;
        private readonly IUiDispatcher _uiDispatcher;

        public CodeMetricsViewModel(
            RubberduckParserState state, 
            ICodeMetricsAnalyst analyst, 
            IVBE vbe,
            IUiDispatcher uiDispatcher)
        {
            _state = state;
            _state.StateChanged += OnStateChanged;

            _analyst = analyst;
            _vbe = vbe;
            _uiDispatcher = uiDispatcher;

            OnPropertyChanged(nameof(Projects));
        }

        private bool _unparsed = true;
        public bool Unparsed
        {
            get => _unparsed;
            set
            {
                if (_unparsed == value)
                {
                    return;
                }
                _unparsed = value;
                OnPropertyChanged();
            }
        }

        private void OnStateChanged(object sender, ParserStateEventArgs e)
        {
            Unparsed = false;
            IsBusy = _state.Status != ParserState.Pending && _state.Status <= ParserState.ResolvedDeclarations;

            if (e.State == ParserState.ResolvedDeclarations)
            {
                Synchronize(_state.DeclarationFinder.AllUserDeclarations);
            }
        }

        private void Synchronize(IEnumerable<Declaration> declarations)
        {
            var metricResults = _analyst.GetMetrics(_state);
            _resultsByDeclaration = metricResults.GroupBy(r => r.Declaration).ToDictionary(g => g.Key, g => g.ToList());

            //We have to wait for the task to guarantee that no new parse starts invalidating all cached components.
            _uiDispatcher.StartTask(() =>
            {
                var updates = declarations.ToList();
                var existing = Projects.OfType<CodeExplorerProjectViewModel>().ToList();

                foreach (var project in existing)
                {
                    project.Synchronize(ref updates);
                    if (project.Declaration is null)
                    {
                        Projects.Remove(project);
                    }
                }

                var adding = updates.OfType<ProjectDeclaration>().ToList();

                foreach (var project in adding)
                {
                    var model = new CodeExplorerProjectViewModel(project, ref updates, _state, _vbe, _state.ProjectsProvider,false);
                    Projects.Add(model);
                }
            }).Wait();
        }

        private ICodeExplorerNode _selectedItem;
        public ICodeExplorerNode SelectedItem
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

        public ObservableCollection<ICodeExplorerNode> Projects { get; } = new ObservableCollection<ICodeExplorerNode>();

        private Dictionary<Declaration, List<ICodeMetricResult>> _resultsByDeclaration;
        public ObservableCollection<ICodeMetricResult> Metrics
        {
            get
            {
                var results = _resultsByDeclaration?.FirstOrDefault(f => ReferenceEquals(f.Key, SelectedItem?.Declaration));
                return results?.Value == null ? new ObservableCollection<ICodeMetricResult>() : new ObservableCollection<ICodeMetricResult>(results.Value.Value);
            }
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                _isBusy = value;
                OnPropertyChanged();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;

        private void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }
            _isDisposed = true;

            _state.StateChanged -= OnStateChanged;
        }
    }
}
