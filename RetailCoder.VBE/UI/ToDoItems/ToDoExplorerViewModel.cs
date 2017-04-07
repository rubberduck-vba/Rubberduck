using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Settings;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.ToDoItems
{
    public sealed class ToDoExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly IGeneralConfigService _configService;
        private readonly IOperatingSystem _operatingSystem;

        public ToDoExplorerViewModel(RubberduckParserState state, IGeneralConfigService configService, IOperatingSystem operatingSystem)
        {
            _state = state;
            _configService = configService;
            _operatingSystem = operatingSystem;
            _state.StateChanged += HandleStateChanged;

            _setMarkerGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByMarker = (bool)param;
                GroupByLocation = !(bool)param;
            });

            _setLocationGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
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

        private readonly CommandBase _setMarkerGroupingCommand;
        public CommandBase SetMarkerGroupingCommand { get { return _setMarkerGroupingCommand; } }

        private readonly CommandBase _setLocationGroupingCommand;
        public CommandBase SetLocationGroupingCommand { get { return _setLocationGroupingCommand; } }

        private CommandBase _refreshCommand;
        public CommandBase RefreshCommand
        {
            get
            {
                if (_refreshCommand != null)
                {
                    return _refreshCommand;
                }
                return _refreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                {
                    _state.OnParseRequested(this);
                },
                _ => _state.IsDirty());
            }
        }

        private void HandleStateChanged(object sender, EventArgs e)
        {
            if (_state.Status != ParserState.ResolvedDeclarations)
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

        private CommandBase _removeCommand;
        public CommandBase RemoveCommand
        {
            get
            {
                if (_removeCommand != null)
                {
                    return _removeCommand;
                }
                return _removeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                {
                    if (_selectedItem == null)
                    {
                        return;
                    }

                    var module = _selectedItem.Selection.QualifiedName.Component.CodeModule;
                    {
                        var oldContent = module.GetLines(_selectedItem.Selection.Selection.StartLine, 1);
                        var newContent = oldContent.Remove(_selectedItem.Selection.Selection.StartColumn - 1);

                        module.ReplaceLine(_selectedItem.Selection.Selection.StartLine, newContent);

                        RefreshCommand.Execute(null);
                    }
                });
            }
        }

        private CommandBase _openTodoSettings;
        public CommandBase OpenTodoSettings
        {
            get
            {
                if (_openTodoSettings != null)
                {
                    return _openTodoSettings;
                }
                return _openTodoSettings = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                {
                    using (var window = new SettingsForm(_configService, _operatingSystem, SettingsViews.TodoSettings))
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
                _state.StateChanged -= HandleStateChanged;
            }
        }
    }
}
