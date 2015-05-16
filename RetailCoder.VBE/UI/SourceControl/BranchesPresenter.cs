using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IBranchesPresenter : IProviderPresenter
    {
        void RefreshView();
    }

    public class BranchesPresenter : IBranchesPresenter
    {
        private readonly IBranchesView _view;
        private readonly ICreateBranchView _createView;
        private readonly IMergeView _mergeView;

        public ISourceControlProvider Provider { get; set; }

        public BranchesPresenter
            (            
                IBranchesView view,
                ICreateBranchView createView,
                IMergeView mergeView,
                ISourceControlProvider provider
            )
            :this(view, createView, mergeView)
        {
            this.Provider = provider;
        }

        public BranchesPresenter
            (
                IBranchesView view,
                ICreateBranchView createView,
                IMergeView mergeView
            )
        {
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
            try
            {
                this.Provider.Checkout(_view.Current);
            }
            catch (SourceControlException ex)
            {
                //todo: find a better way of displaying these errors
                MessageBox.Show(ex.InnerException.Message, ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ~BranchesPresenter()
        {
            _createView.Close();
            _mergeView.Close();
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
            this.Provider.CreateBranch(e.BranchName);
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
