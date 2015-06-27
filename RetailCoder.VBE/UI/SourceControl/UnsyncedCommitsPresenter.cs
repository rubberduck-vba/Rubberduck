using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IUnsyncedCommitsPresenter : IProviderPresenter, IRefreshable
    {
    }

    public class UnsyncedCommitsPresenter : IUnsyncedCommitsPresenter
    {
        private readonly IUnsyncedCommitsView _view;

        public ISourceControlProvider Provider { get; set; }

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
            Provider.Push();
        }

        void OnPull(object sender, EventArgs e)
        {
            Provider.Pull();
        }

        void OnFetch(object sender, EventArgs e)
        {
            Provider.Fetch();
            _view.IncomingCommits = Provider.UnsyncedRemoteCommits;
        }

        void OnSync(object sender, EventArgs e)
        {
            Provider.Pull();
            Provider.Push();
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
