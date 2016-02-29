using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IChangesPresenter : IProviderPresenter, IRefreshable
    {
        void Commit();
    }

    public class ChangesPresenter : ProviderPresenterBase, IChangesPresenter
    {
        private readonly IChangesView _view;

        public ChangesPresenter(IChangesView view)
        {
            _view = view;

            _view.Commit += OnCommit;
            _view.CommitMessageChanged += OnCommitMessageChanged;
            _view.SelectedActionChanged += OnSelectedActionChanged;
        }

        public ChangesPresenter(IChangesView view, ISourceControlProvider provider)
            :this(view)
        {
            Provider = provider;
        }

        private void OnSelectedActionChanged(object sender, EventArgs e)
        {
            _view.CommitEnabled = ShouldEnableCommit();
        }

        private void OnCommitMessageChanged(object sender, EventArgs e)
        {
            _view.CommitEnabled = ShouldEnableCommit();
        }

        private bool ShouldEnableCommit()
        {
            return !string.IsNullOrEmpty(_view.CommitMessage);
        }

        public void RefreshView()
        {
            var fileStats = Provider.Status().ToList();

            _view.IncludedChanges = fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Modified)).ToList();
            _view.UntrackedFiles = fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Untracked)).ToList();

            _view.ExcludedChanges = new List<IFileStatusEntry>();

            _view.CurrentBranch = Provider.CurrentBranch.Name;
        }

        public void Commit()
        {
            var changes = _view.IncludedChanges.Select(c => c.FilePath).ToList();
            if (!changes.Any())
            {
                return;
            }

            try
            {
                Provider.Stage(changes);
                Provider.Commit(_view.CommitMessage);

                if (_view.CommitAction == CommitAction.CommitAndSync)
                {
                    Provider.Pull();
                    Provider.Push();
                }

                if (_view.CommitAction == CommitAction.CommitAndPush)
                {
                    Provider.Push();
                }
            }
            catch(SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }
        }

        private void OnCommit(object sender, EventArgs e)
        {
            Commit();
            _view.CommitMessage = string.Empty;
            RefreshView();
        }
    }
}
