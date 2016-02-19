using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
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
        private readonly Dispatcher _dispatcher;
        private readonly RubberduckParserState _state;
        private readonly IGeneralConfigService _configService;

        public ToDoExplorerViewModel(RubberduckParserState state, IGeneralConfigService configService)
        {
            _dispatcher = Dispatcher.CurrentDispatcher;
            _state = state;
            _configService = configService;
        }

        private readonly ObservableCollection<ToDoItem> _items = new ObservableCollection<ToDoItem>();
        public ObservableCollection<ToDoItem> Items { get { return _items; } } 

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
            _dispatcher.Invoke(() =>
            {
                Items.Clear();
                foreach (var item in GetItems())
                {
                    Items.Add(item);
                }
            });
        }

        private ToDoItem _selectedItem;

        public INavigateSource SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value as ToDoItem; 
                OnPropertyChanged();
            }
        }

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
                    if (_selectedItem == null)
                    {
                        return;
                    }
                    var module = _selectedItem.Selection.QualifiedName.Component.CodeModule;

                    var oldContent = module.Lines[_selectedItem.Selection.Selection.StartLine, 1];
                    var newContent = oldContent.Remove(_selectedItem.Selection.Selection.StartColumn - 1);

                    module.ReplaceLine(_selectedItem.Selection.Selection.StartLine, newContent);

                    RefreshCommand.Execute(null);
                });
            }
        }

        private NavigateCommand _navigateCommand;

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
            var markers = _configService.LoadConfiguration().UserSettings.ToDoListSettings.ToDoMarkers;
            return markers.Where(marker => !string.IsNullOrEmpty(marker.Text)
                                         && comment.CommentText.ToLowerInvariant().Contains(marker.Text.ToLowerInvariant()))
                           .Select(marker => new ToDoItem(marker.Text, comment));
        }

        private IEnumerable<ToDoItem> GetItems()
        {
            return _state.AllComments.SelectMany(GetToDoMarkers);
        }
    }
}
