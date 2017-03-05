using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using NLog;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Controls;
using Rubberduck.UI.Settings;

namespace Rubberduck.UI.Inspections
{
    public sealed class InspectionResultsViewModel : ViewModelBase, INavigateSelection, IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly IInspector _inspector;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;
        private readonly IOperatingSystem _operatingSystem;

        public InspectionResultsViewModel(RubberduckParserState state, IInspector inspector, 
            INavigateCommand navigateCommand, ReparseCommand reparseCommand,
            IClipboardWriter clipboard, IGeneralConfigService configService, IOperatingSystem operatingSystem)
        {
            _state = state;
            _inspector = inspector;
            _navigateCommand = navigateCommand;
            _clipboard = clipboard;
            _configService = configService;
            _operatingSystem = operatingSystem;
            _refreshCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), 
                o => {
                        IsBusy = true;
                        reparseCommand.Execute(o); 
                     },
                o => !IsBusy && reparseCommand.CanExecute(o));

            _disableInspectionCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteDisableInspectionCommand);
            _quickFixCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixCommand, CanExecuteQuickFixCommand);
            _quickFixInModuleCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixInModuleCommand, _ => SelectedItem != null && _state.Status == ParserState.Ready);
            _quickFixInProjectCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteQuickFixInProjectCommand, _ => SelectedItem != null && _state.Status == ParserState.Ready);
            _copyResultsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCopyResultsCommand, CanExecuteCopyResultsCommand);
            _openSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), OpenSettings);

            _configService.SettingsChanged += _configService_SettingsChanged;
            
            // todo: remove I/O work in constructor
            _runInspectionsOnReparse = _configService.LoadConfiguration().UserSettings.CodeInspectionSettings.RunInspectionsOnSuccessfulParse;

            _setInspectionTypeGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByInspectionType = (bool)param;
                GroupByLocation = !(bool)param;
            });

            _setLocationGroupingCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param =>
            {
                GroupByLocation = (bool)param;
                GroupByInspectionType = !(bool)param;
            });

            _state.StateChanged += _state_StateChanged;
        }

        private void _configService_SettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {            
            if (e.InspectionSettingsChanged)
            {
                RefreshInspections();
            }
            _runInspectionsOnReparse = e.RunInspectionsOnReparse;
        }

        private ObservableCollection<IInspectionResult> _results = new ObservableCollection<IInspectionResult>();
        public ObservableCollection<IInspectionResult> Results
        {
            get { return _results; }
            private set
            {
                _results = value;
                OnPropertyChanged();
            }
        }

        private QuickFixBase _defaultFix;

        private INavigateSource _selectedItem;
        public INavigateSource SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value; 
                OnPropertyChanged();

                SelectedInspection = null;
                CanQuickFix = false;
                CanExecuteQuickFixInModule = false;
                CanExecuteQuickFixInProject = false;

                var inspectionResult = _selectedItem as InspectionResultBase;
                if (inspectionResult != null)
                {
                    SelectedInspection = inspectionResult.Inspection;
                    CanQuickFix = inspectionResult.HasQuickFixes;
                    _defaultFix = inspectionResult.DefaultQuickFix;
                    CanExecuteQuickFixInModule = _defaultFix != null && _defaultFix.CanFixInModule;
                }

                CanDisableInspection = SelectedInspection != null;
                CanExecuteQuickFixInProject = _defaultFix != null && _defaultFix.CanFixInProject;
            }
        }

        private IInspection _selectedInspection;
        public IInspection SelectedInspection
        {
            get { return _selectedInspection; }
            set
            {
                _selectedInspection = value;
                OnPropertyChanged();
            }
        }

        private bool _groupByInspectionType = true;
        public bool GroupByInspectionType
        {
            get { return _groupByInspectionType; }
            set
            {
                if (_groupByInspectionType == value) { return; }

                if (value)
                {
                    Results = new ObservableCollection<IInspectionResult>(
                            Results.OrderBy(o => o.Inspection.InspectionType)
                                .ThenBy(t => t.Inspection.Name)
                                .ThenBy(t => t.QualifiedSelection.QualifiedName.Name)
                                .ThenBy(t => t.QualifiedSelection.Selection.StartLine)
                                .ThenBy(t => t.QualifiedSelection.Selection.StartColumn)
                                .ToList());
                }

                _groupByInspectionType = value;
                OnPropertyChanged();
            }
        }

        private bool _groupByLocation;
        public bool GroupByLocation
        {
            get { return _groupByLocation; }
            set
            {
                if (_groupByLocation == value) { return; }

                if (value)
                {
                    Results = new ObservableCollection<IInspectionResult>(
                            Results.OrderBy(o => o.QualifiedSelection.QualifiedName.Name)
                                .ThenBy(t => t.Inspection.Name)
                                .ThenBy(t => t.QualifiedSelection.Selection.StartLine)
                                .ThenBy(t => t.QualifiedSelection.Selection.StartColumn)
                                .ToList());
                }

                _groupByLocation = value;
                OnPropertyChanged();
            }
        }

        private readonly CommandBase _setInspectionTypeGroupingCommand;
        public CommandBase SetInspectionTypeGroupingCommand { get { return _setInspectionTypeGroupingCommand; } }

        private readonly CommandBase _setLocationGroupingCommand;
        public CommandBase SetLocationGroupingCommand { get { return _setLocationGroupingCommand; } }

        private readonly INavigateCommand _navigateCommand;
        public INavigateCommand NavigateCommand { get { return _navigateCommand; } }

        private readonly CommandBase _refreshCommand;
        public CommandBase RefreshCommand { get { return _refreshCommand; } }

        private readonly CommandBase _quickFixCommand;
        public CommandBase QuickFixCommand { get { return _quickFixCommand; } }

        private readonly CommandBase _quickFixInModuleCommand;
        public CommandBase QuickFixInModuleCommand { get { return _quickFixInModuleCommand; } }

        private readonly CommandBase _quickFixInProjectCommand;
        public CommandBase QuickFixInProjectCommand { get { return _quickFixInProjectCommand; } }

        private readonly CommandBase _disableInspectionCommand;
        public CommandBase DisableInspectionCommand { get { return _disableInspectionCommand; } }

        private readonly CommandBase _copyResultsCommand;
        public CommandBase CopyResultsCommand { get { return _copyResultsCommand; } }

        private readonly CommandBase _openSettingsCommand;
        public CommandBase OpenTodoSettings { get { return _openSettingsCommand; } }

        private void OpenSettings(object param)
        {
            using (var window = new SettingsForm(_configService, _operatingSystem, SettingsViews.InspectionSettings))
            {
                window.ShowDialog();
            }
        }

        private bool _canQuickFix;

        public bool CanQuickFix
        {
            get { return _canQuickFix; }
            set
            {
                _canQuickFix = value;
                OnPropertyChanged();
            }
        }

        private bool _isBusy;
        public bool IsBusy { get { return _isBusy; } set { _isBusy = value; OnPropertyChanged(); } }

        private bool _runInspectionsOnReparse;
        private void _state_StateChanged(object sender, EventArgs e)
        {
            if (_state.Status == ParserState.Error || _state.Status == ParserState.ResolverError)
            {
                IsBusy = false;
                return;
            }

            if (_state.Status != ParserState.Ready)
            {
                IsBusy = true;
                return;
            }

            if (sender == this || _runInspectionsOnReparse)
            {
                RefreshInspections();
            }
        }

        private async void RefreshInspections()
        {
            var stopwatch = Stopwatch.StartNew();
            IsBusy = true;

            var results = (await _inspector.FindIssuesAsync(_state, CancellationToken.None)).ToList();
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

            UiDispatcher.Invoke(() =>
            {
                Results = new ObservableCollection<IInspectionResult>(results);

                IsBusy = false;
                SelectedItem = null;
            });

            stopwatch.Stop();
            LogManager.GetCurrentClassLogger().Trace("Inspections loaded in {0}ms", stopwatch.ElapsedMilliseconds);
        }

        private void ExecuteQuickFixes(IEnumerable<QuickFixBase> quickFixes)
        {
            var fixes = quickFixes.ToList();
            var completed = 0;
            var cancelled = 0;
            foreach (var quickFix in fixes)
            {
                quickFix.IsCancelled = false;
                quickFix.Fix();
                completed++;

                if (quickFix.IsCancelled)
                {
                    cancelled++;
                    break;
                }
            }

            // refresh if any quickfix has completed without cancelling:
            if (completed != 0 && cancelled < completed)
            {
                Task.Run(() => _refreshCommand.Execute(null));
            }
        }

        private void ExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as QuickFixBase;
            if (quickFix == null)
            {
                return;
            }

            ExecuteQuickFixes(new[] {quickFix});
        }

        private bool CanExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as QuickFixBase;
            return !IsBusy && quickFix != null && _state.Status == ParserState.Ready;
        }

        private bool _canExecuteQuickFixInModule;
        public bool CanExecuteQuickFixInModule
        {
            get { return _canExecuteQuickFixInModule; }
            set { _canExecuteQuickFixInModule = value; OnPropertyChanged(); }
        }

        private void ExecuteQuickFixInModuleCommand(object parameter)
        {
            if (_defaultFix == null)
            {
                return;
            }

            var selectedResult = SelectedItem as InspectionResultBase;
            if (selectedResult == null)
            {
                return;
            }

            var filteredResults = _results
                .Where(result => result.Inspection == SelectedInspection
                              && result.QualifiedSelection.QualifiedName == selectedResult.QualifiedSelection.QualifiedName)
                .ToList();

            var items = filteredResults
                .Where(result => !(result is AggregateInspectionResult))
                .Select(item => item.QuickFixes.Single(fix => fix.GetType() == _defaultFix.GetType()))
                .Union(filteredResults.OfType<AggregateInspectionResult>().Select(aggregate => aggregate.DefaultQuickFix))
                .OrderByDescending(fix => fix.Selection);
            ExecuteQuickFixes(items);
        }

        private bool _canExecuteQuickFixInProject;
        public bool CanExecuteQuickFixInProject
        {
            get { return _canExecuteQuickFixInProject; }
            set { _canExecuteQuickFixInProject = value; OnPropertyChanged(); }
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

            Task.Run(() => _configService.SaveConfiguration(config)).ContinueWith(t => _refreshCommand.Execute(null));
        }

        private bool _canDisableInspection;

        public bool CanDisableInspection
        {
            get { return _canDisableInspection; }
            set { _canDisableInspection = value; OnPropertyChanged(); }
        }

        private void ExecuteQuickFixInProjectCommand(object parameter)
        {
            if (_defaultFix == null)
            {
                return;
            }

            var items = _results.Where(result => result.Inspection == SelectedInspection)
                .Select(item => item.QuickFixes.Single(fix => fix.GetType() == _defaultFix.GetType()))
                .OrderBy(item => item.Selection.QualifiedName.ComponentName)
                .ThenByDescending(item => item.Selection.Selection.EndLine)
                .ThenByDescending(item => item.Selection.Selection.EndColumn);

            ExecuteQuickFixes(items);
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            const string XML_SPREADSHEET_DATA_FORMAT = "XML Spreadsheet";
            if (_results == null)
            {
                return;
            }
            ColumnInfo[] ColumnInfos = { new ColumnInfo("Type"), new ColumnInfo("Project"), new ColumnInfo("Component"), new ColumnInfo("Issue"), new ColumnInfo("Line", hAlignment.Right), new ColumnInfo("Column", hAlignment.Right) };

            var aResults = _results.Select(result => result.ToArray()).ToArray();

            var resource = _results.Count == 1
                ? RubberduckUI.CodeInspections_NumberOfIssuesFound_Singular
                : RubberduckUI.CodeInspections_NumberOfIssuesFound_Plural;

            var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InvariantCulture), _results.Count);

            var textResults = title + Environment.NewLine + string.Join("", _results.Select(result => result.ToClipboardString() + Environment.NewLine).ToArray());
            var csvResults = ExportFormatter.Csv(aResults, title,ColumnInfos);
            var htmlResults = ExportFormatter.HtmlClipboardFragment(aResults, title,ColumnInfos);
            var rtfResults = ExportFormatter.RTF(aResults, title);

            MemoryStream strm1 = ExportFormatter.XmlSpreadsheetNew(aResults, title, ColumnInfos);
            //Add the formats from richest formatting to least formatting
            _clipboard.AppendStream(DataFormats.GetDataFormat(XML_SPREADSHEET_DATA_FORMAT).Name, strm1);
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

        public void Dispose()
        {
            if (_state != null)
            {
                _state.StateChanged -= _state_StateChanged;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= _configService_SettingsChanged;
            }

            if (_inspector != null)
            {
                _inspector.Dispose();
            }
        }
    }
}
