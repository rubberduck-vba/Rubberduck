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
            if (Provider == null)
            {
                return;
            }

            try
            {
                Provider.Push();

                _view.IncomingCommits = Provider.UnsyncedRemoteCommits;
                _view.OutgoingCommits = Provider.UnsyncedLocalCommits;
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }
        }

        void OnPull(object sender, EventArgs e)
        {
            if (Provider == null)
            {
                return;
            }

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
            FetchCommits();
        }

        private void FetchCommits()
        {
            if (Provider == null)
            {
                return;
            }

            try
            {
                Provider.Fetch();
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }

            _view.IncomingCommits = Provider.UnsyncedRemoteCommits;
            _view.OutgoingCommits = Provider.UnsyncedLocalCommits;
        }

        void OnSync(object sender, EventArgs e)
        {
            if (Provider == null)
            {
                return;
            }

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
            if (Provider != null)
            {
                _view.CurrentBranch = Provider.CurrentBranch.Name;
                FetchCommits();
            }
        }
    }
}
