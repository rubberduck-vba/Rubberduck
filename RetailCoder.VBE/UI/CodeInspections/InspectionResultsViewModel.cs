using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.CodeInspections
{
    public class InspectionResultsViewModel : ViewModelBase
    {
        private readonly IRubberduckParser _parser;
        private readonly IInspector _inspector;
        private readonly VBE _vbe;
        private readonly INavigateCommand _navigateCommand;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;

        public InspectionResultsViewModel(IRubberduckParser parser, IInspector inspector, VBE vbe, INavigateCommand navigateCommand, IClipboardWriter clipboard, IGeneralConfigService configService)
        {
            _parser = parser;
            _inspector = inspector;
            _vbe = vbe;
            _navigateCommand = navigateCommand;
            _clipboard = clipboard;
            _configService = configService;
            _refreshCommand = new DelegateCommand(async param => await Task.Run(() => ExecuteRefreshCommandAsync(param)));
            _disableInspectionCommand = new DelegateCommand(ExecuteDisableInspectionCommand);
            _quickFixCommand = new DelegateCommand(ExecuteQuickFixCommand);
            _quickFixInModuleCommand = new DelegateCommand(ExecuteQuickFixInModuleCommand);
            _quickFixInProjectCommand = new DelegateCommand(ExecuteQuickFixInProjectCommand);
            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand);
        }

        private ObservableCollection<ICodeInspectionResult> _results;

        public ObservableCollection<ICodeInspectionResult> Results
        {
            get { return _results; } 
            set { _results = value; OnPropertyChanged(); }
        }

        private object _selectedItem;
        private CodeInspectionQuickFix _defaultFix;

        public object SelectedItem
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

                var inspectionResult = _selectedItem as CodeInspectionResultBase;

                if (inspectionResult != null)
                {
                    SelectedInspection = inspectionResult.Inspection;
                    CanQuickFix = inspectionResult.HasQuickFixes;
                    _defaultFix = inspectionResult.DefaultQuickFix;
                    CanExecuteQuickFixInModule = _defaultFix != null && _defaultFix.CanFixInModule;
                }
                else
                {
                    var viewGroup = _selectedItem as CollectionViewGroup;
                    if (viewGroup != null)
                    {
                        var grouping = viewGroup;
                        var inspection = grouping.Name as IInspection;
                        if (inspection != null)
                        {
                            SelectedInspection = inspection;
                            var result = _results.FirstOrDefault(item => item.Inspection == inspection);
                            _defaultFix = result == null ? null : result.DefaultQuickFix;
                        }
                    }
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

        public ICommand NavigateCommand { get { return _navigateCommand; } }

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
        public bool CanRefresh { get { return _canRefresh; } private set { _canRefresh = value; OnPropertyChanged(); } }

        private bool _canQuickFix;
        public bool CanQuickFix { get { return _canQuickFix; } set { _canQuickFix = value; OnPropertyChanged(); } }

        private bool _isBusy;
        public bool IsBusy { get { return _isBusy; } set { _isBusy = value; OnPropertyChanged(); } }

        private async void ExecuteRefreshCommandAsync(object parameter)
        {
            CanRefresh = false; // if commands' CanExecute worked as expected, this junk wouldn't be needed
            IsBusy = true;
            var results = await _inspector.FindIssuesAsync(_parser.State, CancellationToken.None);
            Results = new ObservableCollection<ICodeInspectionResult>(results);
            CanRefresh = true;
            IsBusy = false;
            SelectedItem = null;
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

            var selectedResult = SelectedItem as CodeInspectionResultBase;
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
            var results = string.Join("\n", _results.Select(result => result.ToString()), "\n");
            var resource = _results.Count == 1
                ? RubberduckUI.CodeInspections_NumberOfIssuesFound_Singular
                : RubberduckUI.CodeInspections_NumberOfIssuesFound_Plural;
            var text = string.Format(resource, DateTime.Now, _results.Count) + results;

            _clipboard.Write(text);
        }

    }
}
