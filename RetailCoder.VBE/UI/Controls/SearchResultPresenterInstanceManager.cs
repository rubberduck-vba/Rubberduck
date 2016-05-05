using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// A "disposable singleton" factory that creates/returns the same instance to all clients.
    /// </summary>
    public class SearchResultPresenterInstanceManager : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private SearchResultWindow _view;

        public SearchResultPresenterInstanceManager(VBE vbe, AddIn addin)
        {
            _vbe = vbe;
            _addin = addin;
            _view = new SearchResultWindow();
        }

        private SearchResultsDockablePresenter _presenter;
        public SearchResultsDockablePresenter Presenter(ISearchResultsWindowViewModel viewModel)
        {
            if (_presenter == null || _presenter.IsDisposed)
            {
                if (_view.ViewModel == null)
                {
                    _view.ViewModel = viewModel;
                    _view.ViewModel.LastTabClosed += viewModel_LastTabClosed;
                }
                _presenter = new SearchResultsDockablePresenter(_vbe, _addin, _view);
            }

            return _presenter;
        }

        private void viewModel_LastTabClosed(object sender, EventArgs e)
        {
            _presenter.Hide();
        }
        
        public void Dispose()
        {
            if (_view.ViewModel != null)
            {
            _view.ViewModel.LastTabClosed -= viewModel_LastTabClosed;
            }

            if (_presenter != null)
            {
            _presenter.Dispose();
        }
    }
}
}