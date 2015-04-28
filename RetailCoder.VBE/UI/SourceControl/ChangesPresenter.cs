using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IChangesPresenter : IProviderPresenter
    {
        void Refresh();
        void Commit();
    }

    public class ChangesPresenter : IChangesPresenter
    {
        private readonly IChangesView _view;

        public ISourceControlProvider Provider { get; set; }

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
            this.Provider = provider;
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
            return !string.IsNullOrEmpty(_view.CommitMessage) && _view.CommitAction != CommitAction.Unset;
        }

        public void Refresh()
        {
            var fileStats = this.Provider.Status().ToList();

            _view.IncludedChanges = fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Modified)).ToList();
            _view.UntrackedFiles = fileStats.Where(stat => stat.FileStatus.HasFlag(FileStatus.Untracked)).ToList();

            _view.ExcludedChanges = new List<IFileStatusEntry>();
        }

        public void Commit()
        {
            var changes = _view.IncludedChanges.Select(c => c.FilePath).ToList();
            if (!changes.Any())
            {
                return;
            }

            this.Provider.Stage(changes);
            this.Provider.Commit(_view.CommitMessage);

            if (_view.CommitAction == CommitAction.CommitAndSync)
            {
                this.Provider.Pull();
                this.Provider.Push();
            }

            if (_view.CommitAction == CommitAction.CommitAndPush)
            {
                this.Provider.Push();
            }
        }

        private void OnCommit(object sender, EventArgs e)
        {
            Commit();
            _view.CommitMessage = string.Empty;
            Refresh();
        }
    }
}
