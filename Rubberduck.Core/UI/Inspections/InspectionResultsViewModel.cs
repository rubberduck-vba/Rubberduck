﻿using System;
using System.Collections;
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
using System.Windows.Input;
using NLog;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Settings;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Inspections
{
    [Flags]
    public enum InspectionResultsFilter
    {
        None = 0,
        Hint = 1,
        Suggestion = 1 << 1,
        Warning = 1 << 2,
        Error = 1 << 3,
        All = Hint | Suggestion | Warning | Error
    }

    public enum InspectionResultGrouping
    {
        None,
        Type,
        Name,
        Location,
        Severity
    };

    public class DisplayQuickFix
    {
        public IQuickFix Fix { get; }
        public string Description { get; }
        public ICommand Command { get; }

        public DisplayQuickFix(IQuickFix fix, IInspectionResult result, ICommand command)
        {
            Command = command;
            Fix = fix;
            Description = fix.Description(result);
        }
    }

    public sealed class InspectionResultsViewModel : ViewModelBase, INavigateSelection, IComparer<IInspectionResult>, IComparer, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly IInspector _inspector;
        private readonly IQuickFixProvider _quickFixProvider;
        private readonly IClipboardWriter _clipboard;
        private readonly IConfigurationService<Configuration> _configService;
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
            IConfigurationService<Configuration> configService,
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
                    IsBusy = true;
                    _forceRefreshResults = true;
                    var cancellation = new ReparseCancellationFlag();
                    reparseCommand.Execute(cancellation);
                    if (cancellation.Canceled)
                    {
                        IsBusy = false;
                    }
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
            CollapseAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCollapseAll);
            ExpandAllCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteExpandAll);

            _configService.SettingsChanged += _configService_SettingsChanged;
            
            // todo: remove I/O work in constructor
            _runInspectionsOnReparse = _configService.Read().UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse;

            if (CollectionViewSource.GetDefaultView(_results) is ListCollectionView results)
            {
                results.Filter = inspection => InspectionFilter((IInspectionResult)inspection);
                results.CustomSort = this;
                Results = results;
            }

            OnPropertyChanged(nameof(Results));
            Grouping = InspectionResultGrouping.Type;

            _state.StateChanged += HandleStateChanged;
        }

        private void _configService_SettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {            
            if (e.InspectionSettingsChanged)
            {
                _uiDispatcher.Invoke(() =>
                {
                    RefreshCommand.Execute(null);
                });
            }
            _runInspectionsOnReparse = e.RunInspectionsOnReparse;
        }

        private readonly ObservableCollection<IInspectionResult> _results = new ObservableCollection<IInspectionResult>();

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
                SelectedInspection = null;
                CanQuickFix = false;
                CanExecuteQuickFixInProcedure = false;
                CanExecuteQuickFixInModule = false;
                CanExecuteQuickFixInProject = false;

                if (_selectedItem is IInspectionResult inspectionResult)
                {
                    SelectedInspection = inspectionResult.Inspection;

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
            }
        }

        public IEnumerable<DisplayQuickFix> QuickFixes
        {
            get
            {
                if (SelectedItem == null)
                {
                    return Enumerable.Empty<DisplayQuickFix>();
                }

                return _quickFixProvider.QuickFixes(SelectedItem as IInspectionResult)
                    .Select(fix => new DisplayQuickFix(fix, (IInspectionResult)_selectedItem, QuickFixCommand));
            }
        }

        private static readonly Dictionary<InspectionResultGrouping, PropertyGroupDescription> GroupDescriptions = new Dictionary<InspectionResultGrouping, PropertyGroupDescription>
        {
            { InspectionResultGrouping.Type, new PropertyGroupDescription("Inspection", new InspectionTypeConverter()) },
            { InspectionResultGrouping.Name, new PropertyGroupDescription("Inspection.Name") },
            { InspectionResultGrouping.Location, new PropertyGroupDescription("QualifiedSelection.QualifiedName") },
            { InspectionResultGrouping.Severity, new PropertyGroupDescription("Inspection.Severity") }
        };

        private InspectionResultGrouping _grouping;
        public InspectionResultGrouping Grouping
        {
            get => _grouping;
            set
            {
                if (value == _grouping)
                {
                    return;
                }

                _grouping = value;
                // Deferring refresh to avoid a rerendering without grouping
                using (Results.DeferRefresh())
                {
                    Results.GroupDescriptions.Clear();
                    Results.GroupDescriptions.Add(GroupDescriptions[_grouping]);
                }
                OnPropertyChanged();
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

                // updating Filter forces a Refresh
                Results.Filter = i => InspectionFilter((IInspectionResult)i);
            }
        }

        private string _inspectionDescriptionFilter = string.Empty;
        public string InspectionDescriptionFilter
        {
            get => _inspectionDescriptionFilter;
            set
            {
                if (_inspectionDescriptionFilter != value)
                {
                    _inspectionDescriptionFilter = value;
                    OnPropertyChanged();
                    Results.Filter = FilterResults;
                    OnPropertyChanged(nameof(Results));
                }
            }
        }

        private bool FilterResults(object inspectionResult)
        {
            var inspectionResultBase = inspectionResult as InspectionResultBase;
            
            return inspectionResultBase.Description.ToUpper().Contains(InspectionDescriptionFilter.ToUpper()); ;
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
        public CommandBase CollapseAllCommand { get; }
        public CommandBase ExpandAllCommand { get; }

        private void ExecuteCollapseAll(object parameter)
        {
            ExpandedState = false;
        }

        private void ExecuteExpandAll(object parameter)
        {
            ExpandedState = true;
        }

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
            } 
        }

        /// <summary>
        /// A boolean indicating that a local refresh was triggered.
        /// When this is set to true, InspectionResults are refreshed, even when inspecting after successful parsing is disabled.
        /// </summary>
        private bool _forceRefreshResults = false;

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

        private bool _runInspectionsOnReparse;
        private void HandleStateChanged(object sender, ParserStateEventArgs e)
        {
            if (_state.Status == ParserState.Pending || _state.Status == ParserState.Error || _state.Status == ParserState.ResolverError)
            {
                // an error in parser state resets the busy state
                IsBusy = false;
                return;
            }

            if(_state.Status != ParserState.Ready)
            {
                // not an error, but also not finished -> We're busy
                IsBusy = true;
                return;
            }

            if (_state.Status == ParserState.Ready && e.OldState == ParserState.Busy)
            {
                return;
            }

            // push Unparsed to false on the first successful parse
            Unparsed = false;
            if (_runInspectionsOnReparse || _forceRefreshResults)
            {
                RefreshInspections(e.Token);
            }
            else
            {
                //Todo: Find a way to get the actually modified modules in here.
                var modifiedModules = _state.DeclarationFinder.AllModules.ToHashSet();
                InvalidateStaleInspectionResults(modifiedModules);
            }
            IsBusy = false;
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
                IsBusy = false;
                return; //We throw away the partial results.
            }

            stopwatch.Stop();
            Logger.Trace("Inspection results returned in {0}ms", stopwatch.ElapsedMilliseconds);

            _uiDispatcher.Invoke(() =>
            {
                stopwatch = Stopwatch.StartNew();
                try
                {
                    _results.Clear();
                    foreach (var result in results)
                    {
                        _results.Add(result);
                    }
                    Results.Refresh();
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown trying to refresh the inspection results view on the UI thread.");
                }
                finally
                {
                    IsBusy = false;
                    // refreshing results is only disabled when successful
                    // It's basically a "refresh on success once".
                    _forceRefreshResults = false;
                }

                stopwatch.Stop();
                Logger.Trace("Inspection results rendered in {0}ms", stopwatch.ElapsedMilliseconds);
            });
        }

        private void InvalidateUIStaleInspectionResults(ICollection<IInspectionResult> staleResults)
        {
            _uiDispatcher.Invoke(() =>
            {
                foreach (var staleResult in staleResults)
                {
                    _results.Remove(staleResult);
                }
                Results.Refresh();
            });
        }

        private void InvalidateStaleInspectionResults(ICollection<QualifiedModuleName> modifiedModules)
        {
            // materialize the collection to take work off of the UI thread
            var staleResults = _results.Where(result => result.ChangesInvalidateResult(modifiedModules)).ToList();
            InvalidateUIStaleInspectionResults(staleResults);
        }

        private void ExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as IQuickFix;
            _quickFixProvider.Fix(quickFix, SelectedItem as IInspectionResult);
        }

        private bool CanExecuteQuickFixCommand(object parameter)
        {
            return !IsBusy && parameter is IQuickFix && _state.Status == ParserState.Ready;
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

            var config = _configService.Read();

            var setting = config.UserSettings.CodeInspectionSettings.CodeInspections.Single(e => e.Name == _selectedInspection.Name);
            setting.Severity = CodeInspectionSeverity.DoNotShow;

            Task.Run(() => _configService.Save(config));

            // remove inspection results of the selected inspection from the UI
            // collection is materialized to take work off of the UI thread
            InvalidateUIStaleInspectionResults(_results.Where(i => i.Inspection == _selectedInspection).ToList());
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

            if (!(SelectedItem is IInspectionResult selectedResult))
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

            if (!(SelectedItem is IInspectionResult selectedResult))
            {
                return;
            }

            _quickFixProvider.FixAll(_defaultFix, selectedResult.Inspection.GetType(), Results.OfType<IInspectionResult>());
        }

        private static readonly List<(string Name, hAlignment alignment)> ResultColumns = new List<(string Name, hAlignment alignment)>
        {
            (Resources.Inspections.InspectionsUI.ExportColumnHeader_Type, hAlignment.Left),
            (Resources.Inspections.InspectionsUI.ExportColumnHeader_Project, hAlignment.Left),
            (Resources.Inspections.InspectionsUI.ExportColumnHeader_Component, hAlignment.Left),
            (Resources.Inspections.InspectionsUI.ExportColumnHeader_Issue, hAlignment.Left),
            (Resources.Inspections.InspectionsUI.ExportColumnHeader_Line, hAlignment.Right),
            (Resources.Inspections.InspectionsUI.ExportColumnHeader_Column, hAlignment.Right)
        };

        private static readonly ColumnInfo[] ColumnInformation = ResultColumns.Select(column => new ColumnInfo(column.Name, column.alignment)).ToArray();

        private void ExecuteCopyResultsCommand(object parameter)
        {
            const string xmlSpreadsheetDataFormat = "XML Spreadsheet";
            if (Results == null)
            {
                return;
            }

            var resultArray = Results.OfType<IExportable>().Select(result => result.ToArray()).ToArray();

            var resource = resultArray.Count() == 1
                ? Resources.RubberduckUI.CodeInspections_NumberOfIssuesFound_Singular
                : Resources.RubberduckUI.CodeInspections_NumberOfIssuesFound_Plural;

            var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InvariantCulture), resultArray.Count());

            var textResults = title + Environment.NewLine + string.Join(string.Empty, Results.OfType<IExportable>().Select(result => result.ToClipboardString() + Environment.NewLine).ToArray());
            var csvResults = ExportFormatter.Csv(resultArray, title, ColumnInformation);
            var htmlResults = ExportFormatter.HtmlClipboardFragment(resultArray, title, ColumnInformation);
            var rtfResults = ExportFormatter.RTF(resultArray, title);

            // todo: verify that this disposing this stream breaks the xmlSpreadsheetDataFormat
            var stream = ExportFormatter.XmlSpreadsheetNew(resultArray, title, ColumnInformation);
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

        private static readonly Dictionary<InspectionResultGrouping, List<Comparer<IInspectionResult>>> GroupSorts =
            new Dictionary<InspectionResultGrouping, List<Comparer<IInspectionResult>>>
        {
            { InspectionResultGrouping.Type,
                new List<Comparer<IInspectionResult>>
                {
                    InspectionResultComparer.InspectionType,
                    InspectionResultComparer.Location,
                    InspectionResultComparer.Severity,
                    InspectionResultComparer.Name
                }
            },
            { InspectionResultGrouping.Name,
                new List<Comparer<IInspectionResult>>
                {
                    InspectionResultComparer.Name,
                    InspectionResultComparer.Location,
                    InspectionResultComparer.Severity
                }
            },
            { InspectionResultGrouping.Location,
                new List<Comparer<IInspectionResult>>
                {
                    InspectionResultComparer.Location,
                    InspectionResultComparer.Severity,
                    InspectionResultComparer.Name
                }
            },
            { InspectionResultGrouping.Severity,
                new List<Comparer<IInspectionResult>>
                {
                    InspectionResultComparer.Severity,
                    InspectionResultComparer.Location,
                    InspectionResultComparer.Name
                }
            }
        };

        public int Compare(IInspectionResult x, IInspectionResult y)
        {
            return x == y ? 0 : GroupSorts[Grouping].Select(comp => comp.Compare(x, y)).FirstOrDefault(result => result != 0);
        }

        public int Compare(object x, object y)
        {
            if (x == y)
            {
                return 0;
            }

            if (!(x is IInspectionResult first))
            {
                return -1;
            }

            if (!(y is IInspectionResult second))
            {
                return -1;
            }

            return Compare(first, second);
        }

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
