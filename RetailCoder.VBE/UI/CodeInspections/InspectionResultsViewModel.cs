using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.CodeInspections
{
    public class InspectionResultsViewModel : ViewModelBase
    {
        private readonly RubberduckParserState _state;
        private readonly IInspector _inspector;
        private readonly VBE _vbe;
        private readonly IClipboardWriter _clipboard;
        private readonly IGeneralConfigService _configService;

        public InspectionResultsViewModel(RubberduckParserState state, IInspector inspector, VBE vbe, INavigateCommand navigateCommand, IClipboardWriter clipboard, IGeneralConfigService configService)
        {
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

        private readonly INavigateCommand _navigateCommand;
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
            //_state.OnParseRequested();
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
            Results = new ObservableCollection<ICodeInspectionResult>(results);
            CanRefresh = true;
            IsBusy = false;
            SelectedItem = null;

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
            if (_results == null)
            {
                return;
            }
            var aResults = _results.Select(result => result.ToArray()).ToArray();

            var resource = _results.Count == 1
                ? RubberduckUI.CodeInspections_NumberOfIssuesFound_Singular
                : RubberduckUI.CodeInspections_NumberOfIssuesFound_Plural;

            var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InstalledUICulture), _results.Count);

            var csvResults = ExportFormatter.Csv(aResults, title);

//            //13 + 20 + 18 + 24 + 22   :    14 + 20    :   18 + 9 + 7
//            string CfHtmlHeader = "Version:1.0\r\n" +
//                               "StartHTML:0000000105\r\n" +
//                               "EndHTML:{0}\r\n" +
//                               "StartFragment:0000000301\r\n" +
//                               "EndFragment:{1}\r\n" +
//                               "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" +
//                               "<html xmlns=\"http://www.w3.org/1999/xhtml\">\r\n" + 
//                               "" +
//                               "<body>\r\n" +
//                               "<!--StartFragment-->{2}<!--EndFragment-->\r\n" + 
//                               "</body>\r\n" + 
//                               "</html>";

            var textResults = string.Join("", _results.Select(result => result.ToString() + Environment.NewLine).ToArray());
//            var csvResults = string.Join("", _results.Select(result => result.ToCsvString() + Environment.NewLine).ToArray());
//            var htmlResults = string.Join("", _results.Select(result => result.ToHtmlString()).ToArray());
//            var htmlHeader = _results.Select(result => result.ToHtmlHeaderString()).First();

            var text = title + Environment.NewLine + textResults;
            //var csv = "\"" + title + "\"" + Environment.NewLine + csvResults;
//            string html = string.Format("<table cellspacing='0' style='border-bottom: 0.5pt solid #000000;'><tr><td colspan='5'>{0}</td></tr>{1}{2}</table>", title, htmlHeader, htmlResults);

//            long fragmentEnd = 301 + html.Length;
//            long htmlEnd = fragmentEnd + 35;
//            string CfHtml = string.Format(CfHtmlHeader, htmlEnd.ToString("0000000000"), fragmentEnd.ToString("0000000000"), html);
            
            //Add the formats from richest formatting to least formatting
//            _clipboard.AppendData(DataFormats.Html, CfHtml);
            _clipboard.AppendData(DataFormats.CommaSeparatedValue, csvResults);
            _clipboard.AppendData(DataFormats.UnicodeText, text);

            _clipboard.Flush();
        }

        private bool CanExecuteCopyResultsCommand(object parameter)
        {
            return !IsBusy && _results != null && _results.Any();
        }
    }
}
