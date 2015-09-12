using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.CodeInspections
{
    public class InspectionResultsViewModel : ViewModelBase
    {
        private readonly IInspector _inspector;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly VBE _vbe;

        public InspectionResultsViewModel(IInspector inspector, ICodePaneWrapperFactory wrapperFactory, VBE vbe)
        {
            _inspector = inspector;
            _wrapperFactory = wrapperFactory;
            _vbe = vbe;

            _refreshCommand = new DelegateCommand(ExecuteRefreshCommandAsync);
            _quickFixCommand = new DelegateCommand(ExecuteQuickFixCommand);
            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand);
            _exportResultsCommand = new DelegateCommand(ExecuteExportResultsCommand);
        }

        private ObservableCollection<ICodeInspectionResult> _results;
        public ObservableCollection<ICodeInspectionResult> Results { get { return _results; } set { _results = value; OnPropertyChanged(); } }

        private CodeInspectionResultBase _selectedItem;
        public CodeInspectionResultBase SelectedItem { get { return _selectedItem; } set { _selectedItem = value; OnPropertyChanged(); } }

        private readonly ICommand _refreshCommand;
        public ICommand RefreshCommand { get { return _refreshCommand; } }

        private readonly ICommand _quickFixCommand;
        public ICommand QuickFixCommand { get { return _quickFixCommand; } }

        private readonly ICommand _copyResultsCommand;
        public ICommand CopyResultsCommand { get { return _copyResultsCommand; } }

        private readonly ICommand _exportResultsCommand;
        public ICommand ExportResultsCommand { get { return _exportResultsCommand; } }

        private bool _canRefresh = true;
        public bool CanRefresh { get { return _canRefresh; } private set { _canRefresh = value; OnPropertyChanged(); } }

        public bool CanQuickFix { get { return _selectedItem != null && _selectedItem.HasQuickFixes; } }

        private async void ExecuteRefreshCommandAsync(object parameter)
        {
            CanRefresh = false;
            var projectParseResult = await _inspector.Parse(_vbe.ActiveVBProject, this);
            var results = await _inspector.FindIssuesAsync(projectParseResult, CancellationToken.None);
            Results = new ObservableCollection<ICodeInspectionResult>(results);
            CanRefresh = true;
        }

        private void ExecuteQuickFixCommand(object parameter)
        {
            var quickFix = parameter as CodeInspectionQuickFix;
            if (quickFix == null)
            {
                return;
            }

            quickFix.Fix();
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            
        }

        private void ExecuteExportResultsCommand(object parameter)
        {
            
        }
    }
}
