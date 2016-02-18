using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.CodeInspections
{
    public class InspectionResultsViewModel : ViewModelBase, INavigateSelection
    {
        private readonly RubberduckParserState _state;
        private readonly IInspector _inspector;
        private readonly VBE _vbe;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;

        public InspectionResultsViewModel(RubberduckParserState state, IInspector inspector, VBE vbe, INavigateCommand navigateCommand, IClipboardWriter clipboard, IGeneralConfigService configService)
        {
            _dispatcher = Dispatcher.CurrentDispatcher;

            _state = state;
            _inspector = inspector;
            _vbe = vbe;
            _navigateCommand = navigateCommand;
            _clipboard = clipboard;
            _configService = configService;
            _refreshCommand = new DelegateCommand(async param => await Task.Run(() => ExecuteRefreshCommandAsync(param)), CanExecuteRefreshCommand);
            _disableInspectionCommand = new DelegateCommand(ExecuteDisableInspectionCommand);
            _quickFixCommand = new DelegateCommand(ExecuteQuickFixCommand, CanExecuteQuickFixCommand);
            _quickFixInModuleCommand = new DelegateCommand(ExecuteQuickFixInModuleCommand);
            _quickFixInProjectCommand = new DelegateCommand(ExecuteQuickFixInProjectCommand);
            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand, CanExecuteCopyResultsCommand);
        }

        private readonly ObservableCollection<ICodeInspectionResult> _results = new ObservableCollection<ICodeInspectionResult>();
        public ObservableCollection<ICodeInspectionResult> Results
        {
            get { return _results; } 
        }

        private CodeInspectionQuickFix _defaultFix;

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

        private readonly INavigateCommand _navigateCommand;
        public INavigateCommand NavigateCommand { get { return _navigateCommand; } }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

        private readonly ICommand _quickFixCommand;
        public ICommand QuickFixCommand { get { return _quickFixCommand; } }

        private readonly ICommand _quickFixInModuleCommand;
        public ICommand QuickFixInModuleCommand { get { return _quickFixInModuleCommand; } }

        private readonly ICommand _quickFixInProjectCommand;
        public ICommand QuickFixInProjectCommand { get { return _quickFixInProjectCommand; } }

        private readonly ICommand _disableInspectionCommand;
        public ICommand DisableInspectionCommand { get { return _disableInspectionCommand; } }

        private readonly ICommand _copyResultsCommand;
        public ICommand CopyResultsCommand { get { return _copyResultsCommand; } }

        private bool _canRefresh = true;

        public bool CanRefresh
        {
            get { return _canRefresh; }
            private set
            {
                _canRefresh = value; 
                OnPropertyChanged();
            }
        }

        private bool _canQuickFix;
        public bool CanQuickFix { get { return _canQuickFix; } set { _canQuickFix = value; OnPropertyChanged(); } }

        private bool _isBusy;
        public bool IsBusy { get { return _isBusy; } set { _isBusy = value; OnPropertyChanged(); } }

        private async void ExecuteRefreshCommandAsync(object parameter)
        {
            CanRefresh = _vbe.HostApplication() != null;
            if (!CanRefresh)
            {
                return;
            }

            IsBusy = true;

            _state.StateChanged += _state_StateChanged;
            _state.OnParseRequested();
        }

        private bool CanExecuteRefreshCommand(object parameter)
        {
            return !IsBusy;
        }

        private async void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready)
            {
                return;
            }

            var results = await _inspector.FindIssuesAsync(_state, CancellationToken.None);
            _dispatcher.Invoke(() =>
            {
                Results.Clear();
                foreach (var codeInspectionResult in results)
                {
                    Results.Add(codeInspectionResult);
                }
                CanRefresh = true;
                IsBusy = false;
                SelectedItem = null;
            });

            _state.StateChanged -= _state_StateChanged;
        }

        private void ExecuteQuickFixes(IEnumerable<CodeInspectionQuickFix> quickFixes)
        {
            foreach (var quickFix in quickFixes)
            {
                quickFix.Fix();
            }

            Task.Run(() => ExecuteRefreshCommandAsync(null));
        }

        private void ExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as CodeInspectionQuickFix;
            if (quickFix == null)
            {
                return;
            }

            ExecuteQuickFixes(new[] {quickFix});
        }

        private bool CanExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as CodeInspectionQuickFix;
            return !IsBusy && quickFix != null;
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

            var items = _results.Where(result => result.Inspection == SelectedInspection
                && result.QualifiedSelection.QualifiedName == selectedResult.QualifiedSelection.QualifiedName)
                .Select(item => item.QuickFixes.Single(fix => fix.GetType() == _defaultFix.GetType()))
                .OrderByDescending(item => item.Selection.Selection.EndLine)
                .ThenByDescending(item => item.Selection.Selection.EndColumn);

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

            Task.Run(() => _configService.SaveConfiguration(config)).ContinueWith(t => ExecuteRefreshCommandAsync(null));
        }

        private bool _canDisableInspection;
        private readonly Dispatcher _dispatcher;

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
            if (_results == null)
            {
                return;
            }

            var results = string.Join("\n", _results.Select(result => result.ToString() + Environment.NewLine).ToArray());
            var resource = _results.Count == 1
                ? RubberduckUI.CodeInspections_NumberOfIssuesFound_Singular
                : RubberduckUI.CodeInspections_NumberOfIssuesFound_Plural;
            var text = string.Format(resource, DateTime.Now, _results.Count) + Environment.NewLine + results;

            _clipboard.Write(text);
        }

        private bool CanExecuteCopyResultsCommand(object parameter)
        {
            return !IsBusy && _results != null && _results.Any();
        }
    }
}
