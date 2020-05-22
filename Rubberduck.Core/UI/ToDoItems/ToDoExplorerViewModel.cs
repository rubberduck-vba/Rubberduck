using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Text.RegularExpressions;
using System.Windows.Data;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.UI.Command;
using Rubberduck.UI.Settings;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.ToDoExplorer;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.UIContext;
using Rubberduck.SettingsProvider;
using System.Windows.Controls;
using Rubberduck.Formatters;

namespace Rubberduck.UI.ToDoItems
{
    public enum ToDoItemGrouping
    {
        None,
        Marker,
        Location
    };

    public sealed class ToDoExplorerViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly IConfigurationService<Configuration> _configService;
        private readonly ISettingsFormFactory _settingsFormFactory;
        private readonly IUiDispatcher _uiDispatcher;

        public ToDoExplorerViewModel(
            RubberduckParserState state,
            IConfigurationService<Configuration> configService, 
            ISettingsFormFactory settingsFormFactory, 
            IUiDispatcher uiDispatcher,
            INavigateCommand navigateCommand)
        {
            _state = state;
            _configService = configService;
            _settingsFormFactory = settingsFormFactory;
            _uiDispatcher = uiDispatcher;
            _state.StateChanged += HandleStateChanged;

            NavigateCommand = navigateCommand;
            RefreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(),
                _ =>
                {
                    switch(_state.Status)
                    {
                        case ParserState.Ready:
                        case ParserState.Error:
                        case ParserState.ResolverError:
                        case ParserState.UnexpectedError:
                        case ParserState.Pending:
                            _state.OnParseRequested(this);
                            break;
                    }
                },
                _ =>
                {
                    switch (_state.Status)
                    {
                        case ParserState.Ready:
                        case ParserState.Error:
                        case ParserState.ResolverError:
                        case ParserState.UnexpectedError:
                        case ParserState.Pending:
                            return true;
                        default:
                            return false;
                    }
                });
            RemoveCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteRemoveCommand, CanExecuteRemoveCommand);
            CollapseAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCollapseAll);
            ExpandAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteExpandAll);
            CopyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand, CanExecuteCopyResultsCommand);
            OpenTodoSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteOpenTodoSettingsCommand);

            Items = CollectionViewSource.GetDefaultView(_items);
            OnPropertyChanged(nameof(Items));
            Grouping = ToDoItemGrouping.Marker;

            _columnHeaders = _configService.Read().UserSettings.ToDoListSettings.ColumnHeadersInformation;
        }

        private ObservableCollection<ToDoGridViewColumnInfo> _columnHeaders { get; }
        public void UpdateColumnHeaderInformation(ObservableCollection<DataGridColumn> columns)
        {
            _columnHeaders[0].DisplayIndex = columns[0].DisplayIndex;
            _columnHeaders[1].DisplayIndex = columns[1].DisplayIndex;
            _columnHeaders[2].DisplayIndex = columns[2].DisplayIndex;
            _columnHeaders[3].DisplayIndex = columns[3].DisplayIndex;

            _columnHeaders[0].Width = columns[0].Width;
            _columnHeaders[1].Width = columns[1].Width;
            _columnHeaders[2].Width = columns[2].Width;
            _columnHeaders[3].Width = columns[3].Width;

            var userSettings = _configService.Read().UserSettings;
            userSettings.ToDoListSettings.ColumnHeadersInformation = _columnHeaders;
            _configService.Save(new Configuration(userSettings));
        }

        public void UpdateColumnHeaderInformationToMatchCached(ObservableCollection<DataGridColumn> columns)
        {
            columns[0].DisplayIndex = _columnHeaders[0].DisplayIndex;
            columns[1].DisplayIndex = _columnHeaders[1].DisplayIndex;
            columns[2].DisplayIndex = _columnHeaders[2].DisplayIndex;
            columns[3].DisplayIndex = _columnHeaders[3].DisplayIndex;

            columns[0].Width = _columnHeaders[0].Width;
            columns[1].Width = _columnHeaders[1].Width;
            columns[2].Width = _columnHeaders[2].Width;
            columns[3].Width = _columnHeaders[3].Width;
        }

        private readonly ObservableCollection<ToDoItem> _items = new ObservableCollection<ToDoItem>();

        public ICollectionView Items { get; }

        private static readonly Dictionary<ToDoItemGrouping, PropertyGroupDescription> GroupDescriptions = new Dictionary<ToDoItemGrouping, PropertyGroupDescription>
        {
            { ToDoItemGrouping.Marker, new PropertyGroupDescription("Type") },
            { ToDoItemGrouping.Location, new PropertyGroupDescription("Selection.QualifiedName.Name") }
        };

        private ToDoItemGrouping _grouping;
        public ToDoItemGrouping Grouping
        {
            get => _grouping;
            set
            {
                if (value == _grouping)
                {
                    return;
                }

                _grouping = value;
                Items.GroupDescriptions.Clear();
                Items.GroupDescriptions.Add(GroupDescriptions[_grouping]);
                Items.Refresh();
                OnPropertyChanged();
            }
        }

        private bool _expanded;
        public bool ExpandedState
        {
            get => _expanded;
            set
            {
                _expanded = value;
                OnPropertyChanged();
            }
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

        private void HandleStateChanged(object sender, EventArgs e)
        {
            if (_state.Status != ParserState.ResolvedDeclarations)
            {
                return;
            }

            _uiDispatcher.Invoke(() =>
            {
                _items.Clear();
                foreach (var item in _state.AllComments.SelectMany(GetToDoMarkers))
                {
                    _items.Add(item);
                }
            });
        }

        public INavigateCommand NavigateCommand { get; }

        public CommandBase RefreshCommand { get; }

        public CommandBase RemoveCommand { get; }

        public CommandBase CollapseAllCommand { get; }

        public CommandBase ExpandAllCommand { get; }

        public CommandBase CopyResultsCommand { get; }

        public CommandBase OpenTodoSettingsCommand { get; }

        private void ExecuteCollapseAll(object parameter)
        {
            ExpandedState = false;
        }

        private void ExecuteExpandAll(object parameter)
        {
            ExpandedState = true;
        }

        private bool CanExecuteRemoveCommand(object obj) => SelectedItem != null && RefreshCommand.CanExecute(obj);

        private void ExecuteRemoveCommand(object obj)
        {
            if (!CanExecuteRemoveCommand(obj))
            {
                return;
            }

            var component = _state.ProjectsProvider.Component(_selectedItem.Selection.QualifiedName);
            using (var module = component.CodeModule)
            {
                var oldContent = module.GetLines(_selectedItem.Selection.Selection.StartLine, 1);
                var newContent = oldContent.Remove(_selectedItem.Selection.Selection.StartColumn - 1);

                module.ReplaceLine(_selectedItem.Selection.Selection.StartLine, newContent);
            }

            RefreshCommand.Execute(null);
        }

        private bool CanExecuteCopyResultsCommand(object obj) => _items.Any();

        public void ExecuteCopyResultsCommand(object obj)
        {
            const string xmlSpreadsheetDataFormat = "XML Spreadsheet";
            if (!CanExecuteCopyResultsCommand(obj))
            {
                return;
            }

            ColumnInfo[] columnInfos = { new ColumnInfo("Type"), new ColumnInfo("Description"), new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Line", hAlignment.Right), new ColumnInfo("Column", hAlignment.Right) };

            var resultArray = _items
                .Select(item => new ToDoItemFormatter(item))
                .Select(formattedItem => formattedItem.ToArray()).ToArray();

            var resource = _items.Count == 1
                ? ToDoExplorerUI.ToDoExplorer_NumberOfIssuesFound_Singular
                : ToDoExplorerUI.ToDoExplorer_NumberOfIssuesFound_Plural;

            var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InvariantCulture), _items.Count);

            var itemTexts = _items
                .Select(item => new ToDoItemFormatter(item))
                .Select(formattedItem => $"{formattedItem.ToClipboardString()}{Environment.NewLine}")
                .ToArray();
            var textResults = $"{title}{Environment.NewLine}{string.Join(string.Empty, itemTexts)}";
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
        }

        public void ExecuteOpenTodoSettingsCommand(object obj)
        {
            using (var window = _settingsFormFactory.Create(SettingsViews.TodoSettings))
            {
                window.ShowDialog();
                _settingsFormFactory.Release(window);
            }
        }

        private IEnumerable<ToDoItem> GetToDoMarkers(CommentNode comment)
        {
            var markers = _configService.Read().UserSettings.ToDoListSettings.ToDoMarkers;
            return markers.Where(marker => !string.IsNullOrEmpty(marker.Text)
                                         && Regex.IsMatch(comment.CommentText, @"\b" + Regex.Escape(marker.Text) + @"\b", RegexOptions.IgnoreCase))
                           .Select(marker => new ToDoItem(marker.Text, comment));
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
