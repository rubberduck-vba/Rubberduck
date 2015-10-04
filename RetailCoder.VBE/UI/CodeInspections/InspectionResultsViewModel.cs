using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeInspections
{
    public class InspectionResultsViewModel : ViewModelBase
    {
        private readonly IInspector _inspector;
        private readonly VBE _vbe;
        private readonly INavigateCommand _navigateCommand;
        private readonly IClipboardWriter _clipboard;

        public InspectionResultsViewModel(IInspector inspector, VBE vbe, INavigateCommand navigateCommand, IClipboardWriter clipboard)
        {
            _inspector = inspector;
            _vbe = vbe;
            _navigateCommand = navigateCommand;
            _clipboard = clipboard;
            _refreshCommand = new DelegateCommand(async param => await Task.Run(() => ExecuteRefreshCommandAsync(param)));
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

        private CodeInspectionResultBase _selectedItem;

        public CodeInspectionResultBase SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value; 
                OnPropertyChanged();
                CanQuickFix = _selectedItem != null && _selectedItem.HasQuickFixes;

                var defaultFix = _selectedItem != null ? _selectedItem.DefaultQuickFix : null;
                CanExecuteQuickFixInModule = defaultFix != null && defaultFix.CanFixInModule;
                CanExecuteQuickFixInProject = defaultFix != null && defaultFix.CanFixInProject;
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

        private readonly ICommand _copyResultsCommand;
        public ICommand CopyResultsCommand { get { return _copyResultsCommand; } }

        private readonly ICommand _exportResultsCommand;
        public ICommand ExportResultsCommand { get { return _exportResultsCommand; } }

        private bool _canRefresh = true;
        public bool CanRefresh { get { return _canRefresh; } private set { _canRefresh = value; OnPropertyChanged(); } }

        private bool _canQuickFix;
        public bool CanQuickFix { get { return _canQuickFix; } set { _canQuickFix = value; OnPropertyChanged(); } }

        private async void ExecuteRefreshCommandAsync(object parameter)
        {
            CanRefresh = false;
            var projectParseResult = await _inspector.Parse(_vbe.ActiveVBProject, this);
            var results = await _inspector.FindIssuesAsync(projectParseResult, CancellationToken.None);
            Results = new ObservableCollection<ICodeInspectionResult>(results);
            CanRefresh = true;
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
            var quickFix = parameter as CodeInspectionQuickFix;
            if (quickFix == null)
            {
                return;
            }

            var items = _results.Where(result => result.Inspection == SelectedItem.Inspection
                && result.QualifiedSelection.QualifiedName == SelectedItem.QualifiedSelection.QualifiedName)
                .Select(item => item.QuickFixes.Single(fix => fix.GetType() == quickFix.GetType()))
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

        private void ExecuteQuickFixInProjectCommand(object parameter)
        {
            var quickFix = parameter as CodeInspectionQuickFix;
            if (quickFix == null)
            {
                return;
            }

            var items = _results.Where(result => result.Inspection == SelectedItem.Inspection
                && result.QualifiedSelection.QualifiedName.Project == SelectedItem.QualifiedSelection.QualifiedName.Project)
                .Select(item => item.QuickFixes.Single(fix => fix.GetType() == quickFix.GetType()))
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
