using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Windows.Input;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Command;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerViewModel : ViewModelBase
    {
        private readonly RubberduckParserState _state;

        public CodeExplorerViewModel(RubberduckParserState state, INavigateCommand navigateCommand)
        {
            _state = state;
            _navigateCommand = navigateCommand;
            _state.StateChanged += ParserState_StateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand);
        }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

        private readonly INavigateCommand _navigateCommand;
        public ICommand NavigateCommand { get { return _navigateCommand; } }

        private object _selectedItem;
        public object SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value; 
                OnPropertyChanged();
            }
        }

        private bool _isBusy;

        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value; 
                OnPropertyChanged();
                CanRefresh = !_isBusy;
            }
        }

        private bool _canRefresh = true;
        public bool CanRefresh
        {
            get { return true /*_canRefresh*/; }
            private set
            {
                _canRefresh = value;
                OnPropertyChanged();
            }
        }

        private ObservableCollection<CodeExplorerProjectViewModel> _projects;
        public ObservableCollection<CodeExplorerProjectViewModel> Projects
        {
            get { return _projects; }
            set
            {
                _projects = value; 
                OnPropertyChanged();
            }
        }

        private void ParserState_StateChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("CodeExplorerViewModel handles StateChanged...");
            IsBusy = _state.Status == ParserState.Parsing;
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            Debug.WriteLine("Creating Code Explorer model...");
            var userDeclarations = _state.AllUserDeclarations
                .GroupBy(declaration => declaration.ProjectName)
                .Where(grouping => grouping.Key != null)
                .ToList();

            if (userDeclarations.Any(grouping => grouping.All(declaration => declaration.DeclarationType != DeclarationType.Project)))
            {
                return;
            }

            Projects = new ObservableCollection<CodeExplorerProjectViewModel>(userDeclarations.Select(grouping => 
                new CodeExplorerProjectViewModel(grouping.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Project), grouping)));
        }

        private void ParserState_ModuleStateChanged(object sender, Parsing.ParseProgressEventArgs e)
        {
            // todo: figure out a way to handle error state.
            // the problem is that the _projects collection might not contain our failing module yet.
        }

        private void ExecuteRefreshCommand(object param)
        {
            Debug.WriteLine("CodeExplorerViewModel.ExecuteRefreshCommand - requesting reparse");
            _state.OnParseRequested(this);
        }
    }
}
