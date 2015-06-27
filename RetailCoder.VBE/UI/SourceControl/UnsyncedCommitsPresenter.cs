using System;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IUnsyncedCommitsPresenter : IProviderPresenter, IRefreshable
    {
    }

    public class UnsyncedCommitsPresenter : ProviderPresenterBase, IUnsyncedCommitsPresenter
    {
        private readonly IUnsyncedCommitsView _view;

        public UnsyncedCommitsPresenter(IUnsyncedCommitsView view)
        {
            _view = view;

            _view.Sync += OnSync;
            _view.Fetch += OnFetch;
            _view.Pull += OnPull;
            _view.Push += OnPush;
        }

        void OnPush(object sender, EventArgs e)
        {
            try
            {
                Provider.Push();
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }
        }

        void OnPull(object sender, EventArgs e)
        {
            try
            {
                Provider.Pull();
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }
        }

        void OnFetch(object sender, EventArgs e)
        {
            try
            {
                Provider.Fetch();
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }
            
            _view.IncomingCommits = Provider.UnsyncedRemoteCommits;
        }

        void OnSync(object sender, EventArgs e)
        {
            try
            {
                Provider.Pull();
                Provider.Push();
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }
        }

        public void RefreshView()
        {
            if (this.Provider != null)
            {
                _view.CurrentBranch = this.Provider.CurrentBranch.Name;
                _view.IncomingCommits = this.Provider.UnsyncedRemoteCommits;
                _view.OutgoingCommits = this.Provider.UnsyncedLocalCommits;
            }
        }
    }
}
