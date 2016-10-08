using System;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// A "disposable singleton" factory that creates/returns the same instance to all clients.
    /// </summary>
    public sealed class SearchResultPresenterInstanceManager : IDisposable
    {
        private readonly IVBE _vbe;
        private readonly IAddIn _addin;
        private SearchResultWindow _view;

        public SearchResultPresenterInstanceManager(IVBE vbe, IAddIn addin)
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
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing) { return; }

            if (_view.ViewModel != null)
            {
                _view.ViewModel.LastTabClosed -= viewModel_LastTabClosed;
            }

            if (_presenter != null)
            {
                _presenter.Dispose();
                _presenter = null;
            }
        }
    }
}
