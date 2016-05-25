using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.ToDoItems
{
    public sealed class ToDoExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly IGeneralConfigService _configService;

        public ToDoExplorerViewModel(RubberduckParserState state, IGeneralConfigService configService)
        {
            _state = state;
            _configService = configService;
            _state.StateChanged += _state_StateChanged;

            _setMarkerGroupingCommand = new DelegateCommand(param =>
            {
                GroupByMarker = (bool)param;
                GroupByLocation = !(bool)param;
            });

            _setLocationGroupingCommand = new DelegateCommand(param =>
            {
                GroupByLocation = (bool)param;
                GroupByMarker = !(bool)param;
            });
        }

        private ObservableCollection<ToDoItem> _items = new ObservableCollection<ToDoItem>();
        public ObservableCollection<ToDoItem> Items
        {
            get { return _items; }
            set
            {
                if (_items != value)
                {
                    _items = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _groupByMarker = true;
        public bool GroupByMarker
        {
            get { return _groupByMarker; }
            set
            {
                if (_groupByMarker != value)
                {
                    _groupByMarker = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _groupByLocation;
        public bool GroupByLocation
        {
            get { return _groupByLocation; }
            set
            {
                if (_groupByLocation != value)
                {
                    _groupByLocation = value;
                    OnPropertyChanged();
                }
            }
        }

        private readonly ICommand _setMarkerGroupingCommand;
        public ICommand SetMarkerGroupingCommand { get { return _setMarkerGroupingCommand; } }

        private readonly ICommand _setLocationGroupingCommand;
        public ICommand SetLocationGroupingCommand { get { return _setLocationGroupingCommand; } }

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
                    _state.OnParseRequested(this);
                });
            }
        }

        private void _state_StateChanged(object sender, EventArgs e)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            Items = new ObservableCollection<ToDoItem>(GetItems());
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

        private ICommand _removeCommand;
        public ICommand RemoveCommand
        {
            get
            {
                if (_removeCommand != null)
                {
                    return _removeCommand;
                }
                return _removeCommand = new DelegateCommand(_ =>
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

        private ICommand _openTodoSettings;
        public ICommand OpenTodoSettings
        {
            get
            {
                if (_openTodoSettings != null)
                {
                    return _openTodoSettings;
                }
                return _openTodoSettings = new DelegateCommand(_ =>
                {
                    using (var window = new SettingsForm(_configService, SettingsViews.TodoSettings))
                    {
                        window.ShowDialog();
                    }
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

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }
        }
    }
}
