using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IBranchesPresenter
    {
        void RefreshView();
    }

    public class BranchesPresenter : IBranchesPresenter
    {
        private readonly ISourceControlProvider _provider;
        private readonly IBranchesView _view;
        private readonly ICreateBranchView _createView;
        private readonly IMergeView _mergeView;

        public BranchesPresenter(
            ISourceControlProvider provider,
            IBranchesView view,
            ICreateBranchView createView,
            IMergeView mergeView
            )
        {
            _provider = provider;
            _view = view;
            _createView = createView;
            _mergeView = mergeView;

            _view.CreateBranch += OnShowCreateBranchView;
            _view.Merge += OnShowMerge;
            _view.SelectedBranchChanged += OnSelectedBranchChanged;

            _createView.Confirm += OnCreateBranch;
            _createView.Cancel += OnCreateViewCancel;
            _createView.UserInputTextChanged += OnCreateBranchTextChanged;

            _mergeView.Confirm += OnMerge;
            _mergeView.Cancel += OnCancelMerge;
            _mergeView.MergeStatusChanged += OnMergeStatusChanged;
        }

        private void OnSelectedBranchChanged(object sender, EventArgs e)
        {
            _provider.Checkout(_view.Current);
        }

        ~BranchesPresenter()
        {
            _createView.Close();
            _mergeView.Close();
        }

        public void RefreshView()
        {
            _view.Local = _provider.Branches.Where(b => !b.IsRemote).Select(b => b.Name).ToList();
            _view.Current = _provider.CurrentBranch.Name;

            var publishedBranchNames = GetFriendlyBranchNames(RemoteBranches());

            _view.Published = publishedBranchNames;
            _view.Unpublished = _provider.Branches.Where(b => !b.IsRemote
                                                            && publishedBranchNames.All(p => b.Name != p)
                                                            )
                                                    .Select(b => b.Name)
                                                    .ToList();
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
            return _provider.Branches.Where(b => b.IsRemote && !b.Name.Contains("/HEAD"));
        }

        private void HideCreateBranchView()
        {
            _createView.UserInputText = string.Empty;
            _createView.Hide();
        }

        private void OnShowCreateBranchView(object sender, EventArgs e)
        {
            _createView.Show();
        }

        private void OnCreateBranch(object sender, BranchCreateArgs e)
        {
            HideCreateBranchView();
            _provider.CreateBranch(e.BranchName);
            RefreshView();
        }

        private void OnCreateViewCancel(object sender, EventArgs e)
        {
            HideCreateBranchView();
        }

        private void OnCreateBranchTextChanged(object sender, EventArgs e)
        {
            _createView.OkayButtonEnabled = !string.IsNullOrEmpty(_createView.UserInputText);
        }

        private void OnShowMerge(object sender, EventArgs e)
        {
            var localBranchNames = _view.Local.ToList();
            _mergeView.SourceSelectorData = localBranchNames;
            _mergeView.DestinationSelectorData = localBranchNames;
            _mergeView.SelectedSourceBranch = _provider.CurrentBranch.Name;

            _mergeView.Show();
        }

        private void OnMerge(object sender, EventArgs e)
        {
            try
            {
                var source = _mergeView.SelectedSourceBranch;
                var destination = _mergeView.SelectedDestinationBranch;

                _provider.Merge(source, destination);
                _view.Current = _provider.CurrentBranch.Name;

                _mergeView.StatusText = string.Format("Successfully Merged {0} into {1}", source, destination);
                _mergeView.Status = MergeStatus.Success;

                _mergeView.Hide();
            }
            catch (SourceControlException ex)
            {
                _mergeView.Status = MergeStatus.Failure;
                _mergeView.StatusText = ex.Message + ": " + ex.InnerException.Message;
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
