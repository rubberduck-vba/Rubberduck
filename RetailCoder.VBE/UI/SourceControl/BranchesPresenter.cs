using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IBranchesPresenter : IProviderPresenter, IRefreshable
    {
        event EventHandler<EventArgs> BranchChanged;
    }

    public class BranchesPresenter : ProviderPresenterBase, IBranchesPresenter
    {
        private readonly IBranchesView _view;
        private readonly ICreateBranchView _createView;
        private readonly IDeleteBranchView _deleteView;
        private readonly IMergeView _mergeView;

        public event EventHandler<EventArgs> BranchChanged;

        public BranchesPresenter
            (            
                IBranchesView view,
                ICreateBranchView createView,
                IDeleteBranchView deleteView,
                IMergeView mergeView,
                ISourceControlProvider provider
            )
            :this(view, createView, deleteView, mergeView)
        {
            this.Provider = provider;
        }

        public BranchesPresenter
            (
                IBranchesView view,
                ICreateBranchView createView,
                IDeleteBranchView deleteView,
                IMergeView mergeView
            )
        {
            _view = view;
            _createView = createView;
            _deleteView = deleteView;
            _mergeView = mergeView;

            _view.CreateBranch += OnShowCreateBranchView;
            _view.DeleteBranch += OnShowDeleteBranchView;
            _view.Merge += OnShowMerge;
            _view.SelectedBranchChanged += OnSelectedBranchChanged;

            _createView.Confirm += OnCreateBranch;
            _createView.Cancel += OnCreateViewCancel;
            _createView.UserInputTextChanged += OnCreateBranchTextChanged;

            _deleteView.Confirm += OnDeleteBranch;
            _deleteView.Cancel += OnDeleteViewCancel;
            _deleteView.SelectionChanged += OnDeleteViewSelectionChanged;

            _mergeView.Confirm += OnMerge;
            _mergeView.Cancel += OnCancelMerge;
            _mergeView.MergeStatusChanged += OnMergeStatusChanged;
        }

        private void OnSelectedBranchChanged(object sender, EventArgs e)
        {
            var currentBranch = _view.Current;

            try
            {
                this.Provider.Checkout(currentBranch);
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }

            if (BranchChanged != null)
            {
                BranchChanged(this, EventArgs.Empty);
            }
        }

        ~BranchesPresenter()
        {
            _createView.Close();
            _mergeView.Close();
            _deleteView.Close();
        }

        public void RefreshView()
        {
            _view.SelectedBranchChanged -= OnSelectedBranchChanged;

            _view.Local = this.Provider.Branches.Where(b => !b.IsRemote).Select(b => b.Name).ToList();
            _view.Current = this.Provider.CurrentBranch.Name;

            var publishedBranchNames = GetFriendlyBranchNames(RemoteBranches());

            _view.Published = publishedBranchNames;
            _view.Unpublished = this.Provider.Branches.Where(b => !b.IsRemote
                                                            && publishedBranchNames.All(p => b.Name != p)
                                                            )
                                                    .Select(b => b.Name)
                                                    .ToList();

            _view.SelectedBranchChanged += OnSelectedBranchChanged;
        }

        private static IList<string> GetFriendlyBranchNames(IEnumerable<IBranch> branches)
        {
            return branches.Select(
                                    b => b.Name.Split(new[] { '/' })
                                                .Last()
                                   ).ToList();
        }

        private IEnumerable<IBranch> RemoteBranches()
        {
            return this.Provider.Branches.Where(b => b.IsRemote && !b.Name.Contains("/HEAD"));
        }

        private void OnShowDeleteBranchView(object sender, EventArgs e)
        {
            if (_view.Local == null) { return; }

            _deleteView.Branches = _view.Local;
            _deleteView.Show();
        }

        private void OnDeleteViewCancel(object sender, EventArgs e)
        {
            _deleteView.Hide();
        }

        private void OnDeleteViewSelectionChanged(object sender, BranchDeleteArgs e)
        {
            _deleteView.OkButtonEnabled = e.BranchName != _view.Current;
        }

        private void OnDeleteBranch(object sender, BranchDeleteArgs e)
        {
            _deleteView.Hide();

            try
            {
                Provider.DeleteBranch(e.BranchName);
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }

            RefreshView();
        }

        private void HideCreateBranchView()
        {
            _createView.UserInputText = string.Empty;
            _createView.Hide();
        }

        private void OnShowCreateBranchView(object sender, EventArgs e)
        {
            if (_view.Local == null) { return; }
            _createView.Show();
        }

        private void OnCreateBranch(object sender, BranchCreateArgs e)
        {
            HideCreateBranchView();

            try
            {
                this.Provider.CreateBranch(e.BranchName);
            }
            catch (SourceControlException ex)
            {
                RaiseActionFailedEvent(ex);
            }            
            
            RefreshView();
        }

        private void OnCreateViewCancel(object sender, EventArgs e)
        {
            HideCreateBranchView();
        }

        private void OnCreateBranchTextChanged(object sender, EventArgs e)
        {
            _createView.IsValidBranchName = !string.IsNullOrEmpty(_createView.UserInputText) &&
                                          !_view.Local.Contains(_createView.UserInputText) &&
                                          !_createView.UserInputText.Any(char.IsWhiteSpace);
        }

        private void OnShowMerge(object sender, EventArgs e)
        {
            if (_view.Local == null) { return; }

            var localBranchNames = _view.Local.ToList();
            _mergeView.SourceSelectorData = localBranchNames;
            _mergeView.DestinationSelectorData = localBranchNames;
            _mergeView.SelectedSourceBranch = this.Provider.CurrentBranch.Name;

            _mergeView.Show();
        }

        private void OnMerge(object sender, EventArgs e)
        {
            try
            {
                var source = _mergeView.SelectedSourceBranch;
                var destination = _mergeView.SelectedDestinationBranch;

                this.Provider.Merge(source, destination);
                _view.Current = this.Provider.CurrentBranch.Name;

                _mergeView.StatusText = string.Format(RubberduckUI.SourceControl_SuccessfulMerge, source, destination);
                _mergeView.Status = MergeStatus.Success;

                _mergeView.Hide();
            }
            catch (SourceControlException ex)
            {
                _mergeView.Status = MergeStatus.Failure;
                _mergeView.StatusText = ex.Message + ": " + ex.InnerException.Message;
                //todo: raise action failed event?
            }
        }

        private void OnCancelMerge(object sender, EventArgs e)
        {
            _mergeView.Hide();
        }

        private void OnMergeStatusChanged(object sender, EventArgs e)
        {

            if (_mergeView.Status == MergeStatus.Unknown)
            {
                _mergeView.StatusText = string.Empty;
                _mergeView.StatusTextVisible = false;
            }
            else
            {
                _mergeView.StatusTextVisible = true;
            }

        }
    }

}
