using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;

namespace Rubberduck.UI.ToDoItems
{
    public class ToDoExplorerViewModel : ViewModelBase, INavigateSelection
    {
        private readonly RubberduckParserState _state;
        private readonly IEnumerable<ToDoMarker> _markers;
        private ListCollectionView _toDos; 
        public ToDoExplorerViewModel(RubberduckParserState state, IGeneralConfigService configService)
        {
            _state = state;
            _markers = configService.LoadConfiguration().UserSettings.ToDoListSettings.ToDoMarkers;

            _uiDispatcher = Dispatcher.CurrentDispatcher;
        }

        public ListCollectionView ToDos {
            get
            {
                return _toDos;
            }
            set
            {
                _toDos = value;
                OnPropertyChanged();
            }
        } 

        private ICommand _refreshCommand;
        public ICommand RefreshCommand
        {
            get
            {
                if (_refreshCommand != null)
                {
                    return _refreshCommand;
                }
                return _refreshCommand = new DelegateCommand(_ =>
                {
                    _state.StateChanged += _state_StateChanged;
                    _state.OnParseRequested();
                });
            }
        }

        private async void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Parsed)
            {
                return;
            }
            var results = await GetItems();
            
            _uiDispatcher.Invoke(() =>
            {
                ToDos = new ListCollectionView(results.ToList());
                if (ToDos.GroupDescriptions != null)
                {
                    ToDos.GroupDescriptions.Add(new PropertyGroupDescription("Type"));
                }
            });
        }

        public INavigateSource SelectedItem { get { return SelectedToDo; } set { SelectedToDo = value as ToDoItem; } }
        public ToDoItem SelectedToDo { get; set; }

        private ICommand _clear;
        public ICommand Remove
        {
            get
            {
                if (_clear != null)
                {
                    return _clear;
                }
                return _clear = new DelegateCommand(_ =>
                {
                    if (SelectedToDo == null)
                    {
                        return;
                    }
                    var module = SelectedToDo.GetSelection().QualifiedName.Component.CodeModule;

                    var oldContent = module.Lines[SelectedToDo.LineNumber, 1];
                    var newContent =
                        oldContent.Remove(SelectedToDo.GetSelection().Selection.StartColumn - 1);

                    module.ReplaceLine(SelectedToDo.LineNumber, newContent);

                    RefreshCommand.Execute(null);
                });
            }
        }

        private NavigateCommand _navigateCommand;
        private readonly Dispatcher _uiDispatcher;

        public INavigateCommand NavigateCommand
        {
            get
            {
                if (_navigateCommand != null)
                {
                    return _navigateCommand;
                }
                return _navigateCommand = new NavigateCommand();
            }
        }

        private IEnumerable<ToDoItem> GetToDoMarkers(CommentNode comment)
        {
            var c = _markers.Where(marker => !string.IsNullOrEmpty(marker.Text));
            var t = _markers.Select(marker => marker.Text.ToLowerInvariant());
            var v = comment.Comment.ToLowerInvariant();
            var y = _markers.Where(marker => comment.Comment.ToLowerInvariant().Contains(marker.Text.ToLowerInvariant()));
            var z = _markers.Select(marker => new ToDoItem(marker.Text, marker.Priority, comment)).ToList();

            return _markers.Where(marker => !string.IsNullOrEmpty(marker.Text)
                && comment.Comment.ToLowerInvariant().Contains(marker.Text.ToLowerInvariant()))
                .Select(marker => new ToDoItem(marker.Text, marker.Priority, comment)).ToList();
        }

        private async Task<IOrderedEnumerable<ToDoItem>> GetItems()
        {
            var markers = _state.AllComments.SelectMany(GetToDoMarkers).ToList();
            var sortedItems = markers.OrderByDescending(item => item.Priority)
                                   .ThenBy(item => item.ProjectName)
                                   .ThenBy(item => item.ModuleName)
                                   .ThenBy(item => item.LineNumber);

            return sortedItems;
        }
    }
}
