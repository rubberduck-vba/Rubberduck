using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Settings;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Inspections
{
    [Flags]
    public enum InspectionResultsFilter
    {
        None = 0,
        Hint = 1,
        Suggestion = 1 << 2,
        Warning = 1 << 3,
        Error = 1 << 4,
        All = Hint | Suggestion | Warning | Error
    }

    public enum InspectionResultGrouping
    {
        Type,
        Name,
        Location,
        Severity
    };

    public class DisplayQuickFix
    {
        public IQuickFix Fix { get; }
        public string Description { get; }

        public DisplayQuickFix(IQuickFix fix, IInspectionResult result)
        {
            Fix = fix;
            Description = fix.Description(result);
        }
    }

    public sealed class InspectionResultsViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly IInspector _inspector;
        private readonly IQuickFixProvider _quickFixProvider;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;
        private readonly ISettingsFormFactory _settingsFormFactory;
        private readonly IUiDispatcher _uiDispatcher;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public InspectionResultsViewModel(
            RubberduckParserState state, 
            IInspector inspector, 
            IQuickFixProvider quickFixProvider,
            INavigateCommand navigateCommand, 
            ReparseCommand reparseCommand,
            IClipboardWriter clipboard, 
            IGeneralConfigService configService, 
            ISettingsFormFactory settingsFormFactory,
            IUiDispatcher uiDispatcher)
        {
            _state = state;
            _inspector = inspector;
            _quickFixProvider = quickFixProvider;
            NavigateCommand = navigateCommand;
            _clipboard = clipboard;
            _configService = configService;
            _settingsFormFactory = settingsFormFactory;
            _uiDispatcher = uiDispatcher;

            RefreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(),
                o =>
                {
                    IsRefreshing = true;
                    IsBusy = true;
                    reparseCommand.Execute(o);
                },
                o => !IsBusy && reparseCommand.CanExecute(o));

            DisableInspectionCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteDisableInspectionCommand);
            QuickFixCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixCommand, CanExecuteQuickFixCommand);
            QuickFixInProcedureCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixInProcedureCommand, _ => SelectedItem != null && _state.Status == ParserState.Ready);
            QuickFixInModuleCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixInModuleCommand, _ => SelectedItem != null && _state.Status == ParserState.Ready);
            QuickFixInProjectCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixInProjectCommand, _ => SelectedItem != null && _state.Status == ParserState.Ready);
            QuickFixInAllProjectsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixInAllProjectsCommand, _ => SelectedItem != null && _state.Status == ParserState.Ready);
            CopyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand, CanExecuteCopyResultsCommand);
            OpenInspectionSettings = new DelegateCommand(LogManager.GetCurrentClassLogger(), OpenSettings);

            _configService.SettingsChanged += _configService_SettingsChanged;
            
            // todo: remove I/O work in constructor
            _runInspectionsOnReparse = _configService.LoadConfiguration().UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse;

            Results = CollectionViewSource.GetDefaultView(_results) as ListCollectionView;
            Results.Filter = inspection => InspectionFilter((IInspectionResult)inspection);          
            OnPropertyChanged(nameof(Results));

            GroupByInspectionType = true;

            _state.StateChanged += HandleStateChanged;
        }

        /// <summary>
        /// Gets/sets a flag indicating whether the parser state changes are a result of our RefreshCommand.
        /// </summary>
        private bool IsRefreshing { get; set; }

        private void _configService_SettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {            
            if (e.InspectionSettingsChanged)
            {
                RefreshCommand.Execute(null);
            }
            _runInspectionsOnReparse = e.RunInspectionsOnReparse;
        }

        private ObservableCollection<IInspectionResult> _results = new ObservableCollection<IInspectionResult>();

        public ICollectionView Results { get; }

        private IQuickFix _defaultFix;

        private INavigateSource _selectedItem;
        public INavigateSource SelectedItem
        {
            get => _selectedItem;
            set
            {
                _selectedItem = value; 
                OnPropertyChanged();
                OnPropertyChanged(nameof(QuickFixes));
                OnPropertyChanged(nameof(SelectedDescription));
                OnPropertyChanged(nameof(SelectedMeta));
                OnPropertyChanged(nameof(SelectedSeverity));

                SelectedInspection = null;
                CanQuickFix = false;
                CanExecuteQuickFixInProcedure = false;
                CanExecuteQuickFixInModule = false;
                CanExecuteQuickFixInProject = false;

                if (_selectedItem is IInspectionResult inspectionResult)
                {
                    SelectedInspection = inspectionResult.Inspection;
                    SelectedSelection = inspectionResult.QualifiedSelection;
                    
                    CanQuickFix = _quickFixProvider.HasQuickFixes(inspectionResult);
                    _defaultFix = _quickFixProvider.QuickFixes(inspectionResult).FirstOrDefault();
                    CanExecuteQuickFixInProcedure = _defaultFix != null && _defaultFix.CanFixInProcedure;
                    CanExecuteQuickFixInModule = _defaultFix != null && _defaultFix.CanFixInModule;
                    CanExecuteQuickFixInModule = _defaultFix != null && _defaultFix.CanFixInProcedure;
                    CanExecuteQuickFixInProject = _defaultFix != null && _defaultFix.CanFixInProject;
                }

                CanDisableInspection = SelectedInspection != null;
            }
        }

        private IInspection _selectedInspection;
        public IInspection SelectedInspection
        {
            get => _selectedInspection;
            set
            {
                _selectedInspection = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SelectedSelection));
            }
        }

        public string SelectedDescription => SelectedInspection?.Description ?? string.Empty;

        public string SelectedMeta => SelectedInspection?.Meta ?? string.Empty;

        public CodeInspectionSeverity SelectedSeverity => SelectedInspection?.Severity ?? CodeInspectionSeverity.DoNotShow;

        public QualifiedSelection SelectedSelection { get; private set; }

        public IEnumerable<DisplayQuickFix> QuickFixes
        {
            get
            {
                if (SelectedItem == null)
                {
                    return Enumerable.Empty<DisplayQuickFix>();
                }

                return _quickFixProvider.QuickFixes(SelectedItem as IInspectionResult)
                    .Select(fix => new DisplayQuickFix(fix, (IInspectionResult) _selectedItem));
            }
        }

        private static readonly Dictionary<InspectionResultGrouping, PropertyGroupDescription> GroupDescriptions = new Dictionary<InspectionResultGrouping, PropertyGroupDescription>
        {
            { InspectionResultGrouping.Type, new PropertyGroupDescription("Inspection", new InspectionTypeConverter()) },
            { InspectionResultGrouping.Name, new PropertyGroupDescription("Inspection.Name") },
            { InspectionResultGrouping.Location, new PropertyGroupDescription("QualifiedSelection.QualifiedName") },
            { InspectionResultGrouping.Severity, new PropertyGroupDescription("Inspection.Severity") }
        };

        private void SetGrouping(InspectionResultGrouping grouping)
        {
            Results.GroupDescriptions.Clear();
            Results.GroupDescriptions.Add(GroupDescriptions[grouping]);
            Results.Refresh();

            OnPropertyChanged(nameof(GroupByInspectionType));
            OnPropertyChanged(nameof(GroupByName));
            OnPropertyChanged(nameof(GroupByLocation));
            OnPropertyChanged(nameof(GroupBySeverity));
        }

        private bool _groupByInspection;
        public bool GroupByInspectionType
        {
            get => _groupByInspection;
            set
            {
                if (value == _groupByInspection)
                {
                    return;
                }

                _groupByInspection = value;
                if (_groupByInspection)
                {
                    _groupByName = false;
                    _groupByLocation = false;
                    _groupBySeverity = false;
                    SetGrouping(InspectionResultGrouping.Type);
                }
            }
        }

        private bool _groupByName;
        public bool GroupByName
        {
            get => _groupByName;
            set
            {
                if (value == _groupByName)
                {
                    return;
                }

                _groupByName = value;
                if (_groupByName)
                {
                    _groupByInspection = false;
                    _groupByLocation = false;
                    _groupBySeverity = false;
                    SetGrouping(InspectionResultGrouping.Name);
                }
            }
        }

        private bool _groupByLocation;
        public bool GroupByLocation
        {
            get => _groupByLocation;
            set
            {
                if (value == _groupByLocation)
                {
                    return;
                }

                _groupByLocation = value;
                if (_groupByLocation)
                {
                    _groupByInspection = false;
                    _groupByName = false;
                    _groupBySeverity = false;
                    SetGrouping(InspectionResultGrouping.Location);
                }
            }
        }

        private bool _groupBySeverity;
        public bool GroupBySeverity
        {
            get => _groupBySeverity;
            set
            {
                if (value == _groupBySeverity)
                {
                    return;
                }

                _groupBySeverity = value;
                if (_groupBySeverity)
                {
                    _groupByInspection = false;
                    _groupByName = false;
                    _groupByLocation = false;
                    SetGrouping(InspectionResultGrouping.Severity);
                }
            }
        }

        private InspectionResultsFilter _filters = InspectionResultsFilter.All;
        public InspectionResultsFilter SelectedFilters
        {
            get => _filters;
            set
            {
                if (value == _filters)
                {
                    return;
                }

                _filters = value;
                OnPropertyChanged();
                Results.Refresh();
            }
        }

        private bool InspectionFilter(IInspectionResult result)
        {
            switch (result.Inspection.Severity)
            {
                case CodeInspectionSeverity.DoNotShow:
                    return false;
                case CodeInspectionSeverity.Hint:
                    return SelectedFilters.HasFlag(InspectionResultsFilter.Hint);
                case CodeInspectionSeverity.Suggestion:
                    return SelectedFilters.HasFlag(InspectionResultsFilter.Suggestion);
                case CodeInspectionSeverity.Warning:
                    return SelectedFilters.HasFlag(InspectionResultsFilter.Warning);
                case CodeInspectionSeverity.Error:
                    return SelectedFilters.HasFlag(InspectionResultsFilter.Error);
                default:
                    return true;    // Not in the enum...
            }     
        }

        public INavigateCommand NavigateCommand { get; }
        public CommandBase RefreshCommand { get; }
        public CommandBase QuickFixCommand { get; }
        public CommandBase QuickFixInProcedureCommand { get; }
        public CommandBase QuickFixInModuleCommand { get; }
        public CommandBase QuickFixInProjectCommand { get; }
        public CommandBase QuickFixInAllProjectsCommand { get; }
        public CommandBase DisableInspectionCommand { get; }
        public CommandBase CopyResultsCommand { get; }
        public CommandBase OpenInspectionSettings { get; }

        private void OpenSettings(object param)
        {
            using (var window = _settingsFormFactory.Create(SettingsViews.InspectionSettings))
            {
                window.ShowDialog();
                _settingsFormFactory.Release(window);
            }
        }

        private bool _canQuickFix;

        public bool CanQuickFix
        {
            get => _canQuickFix;
            set
            {
                _canQuickFix = value;
                OnPropertyChanged();
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
                OnPropertyChanged("EmptyUIRefreshMessageVisibility");
            } 
        }

        private bool _runInspectionsOnReparse;
        private void HandleStateChanged(object sender, ParserStateEventArgs e)
        {
            if(!IsRefreshing && (_state.Status == ParserState.Pending || _state.Status == ParserState.Error || _state.Status == ParserState.ResolverError))
            {
                IsBusy = false;
                return;
            }

            if(_state.Status != ParserState.Ready)
            {
                return;
            }

            if (_state.Status == ParserState.Ready && e.OldState == ParserState.Busy)
            {
                return;
            }

            if (_runInspectionsOnReparse || IsRefreshing)
            {
                RefreshInspections(e.Token);
            }
            else
            {
                //Todo: Find a way to get the actually modified modules in here.
                var modifiedModules = _state.DeclarationFinder.AllModules.ToHashSet();
                InvalidateStaleInspectionResults(modifiedModules);
            }
        }

        private async void RefreshInspections(CancellationToken token)
        {
            var stopwatch = Stopwatch.StartNew();
            IsBusy = true;

            List<IInspectionResult> results;
            try
            {
                var inspectionResults = await _inspector.FindIssuesAsync(_state, token);
                results = inspectionResults.ToList();
            }
            catch (OperationCanceledException)
            {
                Logger.Debug("Inspections got canceled.");
                return; //We throw away the partial results.
            }

            if (GroupByInspectionType)
            {
                results = results.OrderBy(o => o.Inspection.InspectionType)
                    .ThenBy(t => t.Inspection.Name)
                    .ThenBy(t => t.QualifiedSelection.QualifiedName.Name)
                    .ThenBy(t => t.QualifiedSelection.Selection.StartLine)
                    .ThenBy(t => t.QualifiedSelection.Selection.StartColumn)
                    .ToList();
            }
            else
            {
                results = results.OrderBy(o => o.QualifiedSelection.QualifiedName.Name)
                    .ThenBy(t => t.Inspection.Name)
                    .ThenBy(t => t.QualifiedSelection.Selection.StartLine)
                    .ThenBy(t => t.QualifiedSelection.Selection.StartColumn)
                    .ToList();
            }

            _uiDispatcher.Invoke(() =>
            {
                _results.Clear();
                foreach (var result in results)
                {
                    _results.Add(result);
                }

                Results.Refresh();

                try
                {
                    IsBusy = false;
                    OnPropertyChanged("EmptyUIRefreshVisibility");
                    IsRefreshing = false;
                    SelectedItem = null;
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown trying to refresh the inspection results view on th UI thread.");
                }
            });

            stopwatch.Stop();
            LogManager.GetCurrentClassLogger().Trace("Inspections loaded in {0}ms", stopwatch.ElapsedMilliseconds);
        }

        private void InvalidateStaleInspectionResults(ICollection<QualifiedModuleName> modifiedModules)
        {
            var staleResults = Results.Where(result => result.ChangesInvalidateResult(modifiedModules)).ToList();
            _uiDispatcher.Invoke(() =>
            {
                foreach (var staleResult in staleResults)
                {
                    Results.Remove(staleResult);
                }
            });
        }

        private void ExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as IQuickFix;
            _quickFixProvider.Fix(quickFix, SelectedItem as IInspectionResult);
        }

        private bool CanExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as IQuickFix;
            return !IsBusy && quickFix != null && _state.Status == ParserState.Ready;
        }

        private bool _canExecuteQuickFixInProcedure;
        public bool CanExecuteQuickFixInProcedure
        {
            get => _canExecuteQuickFixInProcedure;
            set
            {
                _canExecuteQuickFixInProcedure = value;
                OnPropertyChanged();
            }
        }

        private void ExecuteQuickFixInProcedureCommand(object parameter)
        {
            if (_defaultFix == null)
            {
                return;
            }

            var selectedResult = SelectedItem as IInspectionResult;
            if (selectedResult == null)
            {
                return;
            }

            _quickFixProvider.FixInProcedure(_defaultFix, selectedResult.QualifiedMemberName,
                selectedResult.Inspection.GetType(), Results.OfType<IInspectionResult>());
        }

        private bool _canExecuteQuickFixInModule;
        public bool CanExecuteQuickFixInModule
        {
            get => _canExecuteQuickFixInModule;
            set
            {
                _canExecuteQuickFixInModule = value;
                OnPropertyChanged();
            }
        }

        private void ExecuteQuickFixInModuleCommand(object parameter)
        {
            if (_defaultFix == null)
            {
                return;
            }

            var selectedResult = SelectedItem as IInspectionResult;
            if (selectedResult == null)
            {
                return;
            }
            
            _quickFixProvider.FixInModule(_defaultFix, selectedResult.QualifiedSelection,
                selectedResult.Inspection.GetType(), Results.OfType<IInspectionResult>());
        }

        private bool _canExecuteQuickFixInProject;
        public bool CanExecuteQuickFixInProject
        {
            get => _canExecuteQuickFixInProject;
            set
            {
                _canExecuteQuickFixInProject = value;
                OnPropertyChanged();
            }
        }

        private void ExecuteDisableInspectionCommand(object parameter)
        {
            if (_selectedInspection == null)
            {
                return;
            }

            var config = _configService.LoadConfiguration();

            var setting = config.UserSettings.CodeInspectionSettings.CodeInspections.Single(e => e.Name == _selectedInspection.Name);
            setting.Severity = CodeInspectionSeverity.DoNotShow;

            Task.Run(() => _configService.SaveConfiguration(config)).ContinueWith(t => RefreshCommand.Execute(null));
        }

        private bool _canDisableInspection;

        public bool CanDisableInspection
        {
            get => _canDisableInspection;
            set
            {
                _canDisableInspection = value;
                OnPropertyChanged();
            }
        }

        private void ExecuteQuickFixInProjectCommand(object parameter)
        {
            if (_defaultFix == null)
            {
                return;
            }

            var selectedResult = SelectedItem as IInspectionResult;
            if (selectedResult == null)
            {
                return;
            }

            _quickFixProvider.FixInProject(_defaultFix, selectedResult.QualifiedSelection,
                selectedResult.Inspection.GetType(), Results.OfType<IInspectionResult>());
        }

        private void ExecuteQuickFixInAllProjectsCommand(object parameter)
        {
            if (_defaultFix == null)
            {
                return;
            }

            var selectedResult = SelectedItem as IInspectionResult;
            if (selectedResult == null)
            {
                return;
            }

            _quickFixProvider.FixAll(_defaultFix, selectedResult.Inspection.GetType(), Results.OfType<IInspectionResult>());
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            const string xmlSpreadsheetDataFormat = "XML Spreadsheet";
            if (_results == null)
            {
                return;
            }
            ColumnInfo[] columnInfos = { new ColumnInfo("Type"), new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Issue"), new ColumnInfo("Line", hAlignment.Right), new ColumnInfo("Column", hAlignment.Right) };

            var resultArray = _results.OfType<IExportable>().Select(result => result.ToArray()).ToArray();

            var resource = _results.Count == 1
                ? Resources.RubberduckUI.CodeInspections_NumberOfIssuesFound_Singular
                : Resources.RubberduckUI.CodeInspections_NumberOfIssuesFound_Plural;

            var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InvariantCulture), _results.Count);

            var textResults = title + Environment.NewLine + string.Join("", _results.OfType<IExportable>().Select(result => result.ToClipboardString() + Environment.NewLine).ToArray());
            var csvResults = ExportFormatter.Csv(resultArray, title,columnInfos);
            var htmlResults = ExportFormatter.HtmlClipboardFragment(resultArray, title,columnInfos);
            var rtfResults = ExportFormatter.RTF(resultArray, title);

            // todo: verify that this disposing this stream breaks the xmlSpreadsheetDataFormat
            var stream = ExportFormatter.XmlSpreadsheetNew(resultArray, title, columnInfos);
            //Add the formats from richest formatting to least formatting
            _clipboard.AppendStream(DataFormats.GetDataFormat(xmlSpreadsheetDataFormat).Name, stream);
            _clipboard.AppendString(DataFormats.Rtf, rtfResults);
            _clipboard.AppendString(DataFormats.Html, htmlResults);
            _clipboard.AppendString(DataFormats.CommaSeparatedValue, csvResults);
            _clipboard.AppendString(DataFormats.UnicodeText, textResults);

            _clipboard.Flush();
        }

        private bool CanExecuteCopyResultsCommand(object parameter)
        {
            return !IsBusy && _results != null && _results.Any();
        }

        public Visibility EmptyUIRefreshVisibility => _state.ProjectsProvider.Projects().Any() ? Visibility.Hidden : Visibility.Visible;

        public Visibility EmptyUIRefreshMessageVisibility => IsBusy ? Visibility.Hidden : Visibility.Visible;

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= HandleStateChanged;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= _configService_SettingsChanged;
            }

            _inspector?.Dispose();
        }
    }
}
