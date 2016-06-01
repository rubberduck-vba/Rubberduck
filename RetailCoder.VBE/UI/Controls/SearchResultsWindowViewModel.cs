using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Rubberduck.Common;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsWindowViewModel : ViewModelBase, ISearchResultsWindowViewModel
    {
        private readonly IClipboardWriter _clipboard;

        private readonly ObservableCollection<SearchResultsViewModel> _tabs = 
            new ObservableCollection<SearchResultsViewModel>();

        public SearchResultsWindowViewModel(IClipboardWriter clipboard)
        {
            _clipboard = clipboard;
            _copyResultsCommand = new DelegateCommand(ExecuteCopyResultsCommand, CanExecuteCopyResultsCommand);
        }

        public void AddTab(SearchResultsViewModel viewModel)
        {
            viewModel.Close += viewModel_Close;
            _tabs.Add(viewModel);
        }

        void viewModel_Close(object sender, EventArgs e)
        {
            RemoveTab(sender as SearchResultsViewModel);
        }

        private readonly ICommand _copyResultsCommand;
        public ICommand CopyResultsCommand { get { return _copyResultsCommand; } }

        private bool CanExecuteCopyResultsCommand(object parameter)
        {
            return _selectedTab != null;
        }

        private void ExecuteCopyResultsCommand(object parameter)
        {
            const string xmlSpreadsheetDataFormat = "XML Spreadsheet";
            if (_selectedTab == null)
            {
                return;
            }
            //todo: implement in its own class, move .ToArray() and .ToCsvString() to its own interface

            //ColumnInfo[] columnInfos =
            //{
            //    new ColumnInfo("Type"), 
            //    new ColumnInfo("Project"), 
            //    new ColumnInfo("Component"), 
            //    new ColumnInfo("Issue"), 
            //    new ColumnInfo("Line", hAlignment.Right), 
            //    new ColumnInfo("Column", hAlignment.Right)
            //};

            //var items = _selectedTab.SearchResults.ToArray();

            //var title = string.Format("{0} ({1})", _selectedTab.Header, items.Length);

            //var textResults = title + Environment.NewLine + string.Join(string.Empty, items.Select(result => result.ToString() + Environment.NewLine).ToArray());
            //var csvResults = ExportFormatter.Csv(items, title, columnInfos);
            //var htmlResults = ExportFormatter.HtmlClipboardFragment(items, title, columnInfos);
            //var rtfResults = ExportFormatter.RTF(items, title);

            //MemoryStream strm1 = ExportFormatter.XmlSpreadsheetNew(items, title, columnInfos);
            ////Add the formats from richest formatting to least formatting
            //_clipboard.AppendStream(DataFormats.GetDataFormat(xmlSpreadsheetDataFormat).Name, strm1);
            //_clipboard.AppendString(DataFormats.Rtf, rtfResults);
            //_clipboard.AppendString(DataFormats.Html, htmlResults);
            //_clipboard.AppendString(DataFormats.CommaSeparatedValue, csvResults);
            //_clipboard.AppendString(DataFormats.UnicodeText, textResults);

            //_clipboard.Flush();
        }

        public IEnumerable<SearchResultsViewModel> Tabs { get { return _tabs; } }

        private SearchResultsViewModel _selectedTab;

        public SearchResultsViewModel SelectedTab
        {
            get { return _selectedTab; }
            set
            {
                if (_selectedTab != value)
                {
                    _selectedTab = value;
                    OnPropertyChanged();
                }
            }
        }

        private void RemoveTab(SearchResultsViewModel viewModel)
        {
            if (viewModel != null)
            {
                _tabs.Remove(viewModel);
            }

            if (!_tabs.Any())
            {
                OnLastTabClosed();
            }
        }

        public event EventHandler LastTabClosed;
        private void OnLastTabClosed()
        {
            var handler = LastTabClosed;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }
    }
}
