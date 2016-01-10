using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerViewModel : ViewModelBase
    {
        private readonly RubberduckParserState _state;

        public CodeExplorerViewModel(RubberduckParserState state)
        {
            _state = state;
            _state.StateChanged += ParserState_StateChanged;
            _state.ModuleStateChanged += ParserState_ModuleStateChanged;

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommand);
        }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

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

        private IEnumerable<CodeExplorerProjectViewModel> _projects;
        public IEnumerable<CodeExplorerProjectViewModel> Projects
        {
            get { return _projects; }
            set
            {
                _projects = value; 
                OnPropertyChanged();
            }
        }

        private void ParserState_StateChanged(object sender, ParserStateEventArgs e)
        {
            IsBusy = e.State == ParserState.Parsing;
            if (e.State != ParserState.Resolving) // Parsed state is too volatile
            {
                return;
            }

            var userDeclarations = _state.AllUserDeclarations
                .GroupBy(declaration => declaration.Project)
                .ToList();

            Projects = userDeclarations.Select(grouping => 
                new CodeExplorerProjectViewModel(grouping.Single(declaration => declaration.DeclarationType == DeclarationType.Project), grouping));
        }

        private void ParserState_ModuleStateChanged(object sender, Parsing.ParseProgressEventArgs e)
        {
            // todo: figure out a way to handle error state.
            // the problem is that the _projects collection might not contain our failing module yet.
        }

        private void ExecuteRefreshCommand(object param)
        {
            _state.OnParseRequested();
        }
    }
}
