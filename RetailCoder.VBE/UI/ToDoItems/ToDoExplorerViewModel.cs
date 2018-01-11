using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Text.RegularExpressions;
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

            SetMarkerGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByMarker = (bool)param;
                GroupByLocation = !(bool)param;
            });

            SetLocationGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByLocation = (bool)param;
                GroupByMarker = !(bool)param;
            });
        }

        private ObservableCollection<ToDoItem> _items = new ObservableCollection<ToDoItem>();
        public ObservableCollection<ToDoItem> Items
        {
            get => _items;
            set
            {
                if (_items == value)
                {
                    return;
                }

                _items = value;
                OnPropertyChanged();
            }
        }

        private bool _groupByMarker = true;
        public bool GroupByMarker
        {
            get => _groupByMarker;
            set
            {
                if (_groupByMarker == value)
                {
                    return;
                }

                _groupByMarker = value;
                OnPropertyChanged();
            }
        }

        private bool _groupByLocation;
        public bool GroupByLocation
        {
            get => _groupByLocation;
            set
            {
                if (_groupByLocation == value)
                {
                    return;
                }

                _groupByLocation = value;
                OnPropertyChanged();
            }
        }

        public CommandBase SetMarkerGroupingCommand { get; }

        public CommandBase SetLocationGroupingCommand { get; }

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
            get => _selectedItem;
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

                    using (var module = _selectedItem.Selection.QualifiedName.Component.CodeModule)
                    {
                        var oldContent = module.GetLines(_selectedItem.Selection.Selection.StartLine, 1);
                        var newContent = oldContent.Remove(_selectedItem.Selection.Selection.StartColumn - 1);

                        module.ReplaceLine(_selectedItem.Selection.Selection.StartLine, newContent);
                    }
                    RefreshCommand.Execute(null);
                }
                );
            }
        }

        private CommandBase _copyResultsCommand;
        public CommandBase CopyResultsCommand
        {
            get
            {
                if (_copyResultsCommand != null)
                {
                    return _copyResultsCommand;
                }
                return _copyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                {
                    const string xmlSpreadsheetDataFormat = "XML Spreadsheet";
                    if (_items == null)
                    {
                        return;
                    }
                    ColumnInfo[] columnInfos = { new ColumnInfo("Type"), new ColumnInfo("Description"), new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Line", hAlignment.Right), new ColumnInfo("Column", hAlignment.Right) };

                    var resultArray = _items.OfType<IExportable>().Select(result => result.ToArray()).ToArray();

                    var resource = _items.Count == 1
                        ? RubberduckUI.ToDoExplorer_NumberOfIssuesFound_Singular
                        : RubberduckUI.ToDoExplorer_NumberOfIssuesFound_Plural;

                    var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InvariantCulture), _items.Count);

                    var textResults = title + Environment.NewLine + string.Join("", _items.OfType<IExportable>().Select(result => result.ToClipboardString() + Environment.NewLine).ToArray());
                    var csvResults = ExportFormatter.Csv(resultArray, title, columnInfos);
                    var htmlResults = ExportFormatter.HtmlClipboardFragment(resultArray, title, columnInfos);
                    var rtfResults = ExportFormatter.RTF(resultArray, title);

                    // todo: verify that this disposing this stream breaks the xmlSpreadsheetDataFormat
                    var stream = ExportFormatter.XmlSpreadsheetNew(resultArray, title, columnInfos);

                    IClipboardWriter _clipboard = new ClipboardWriter();
                    //Add the formats from richest formatting to least formatting
                    _clipboard.AppendStream(DataFormats.GetDataFormat(xmlSpreadsheetDataFormat).Name, stream);
                    _clipboard.AppendString(DataFormats.Rtf, rtfResults);
                    _clipboard.AppendString(DataFormats.Html, htmlResults);
                    _clipboard.AppendString(DataFormats.CommaSeparatedValue, csvResults);
                    _clipboard.AppendString(DataFormats.UnicodeText, textResults);

                    _clipboard.Flush();

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
                                         && Regex.IsMatch(comment.CommentText, @"\b" + Regex.Escape(marker.Text) + @"\b", RegexOptions.IgnoreCase))
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
