using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class BranchesPresenter
    {
        private readonly ISourceControlProvider _provider;
        private readonly IBranchesView _view;
        private readonly ICreateBranchView _createView;

        public BranchesPresenter(ISourceControlProvider provider, IBranchesView view, ICreateBranchView createView)
        {
            _provider = provider;
            _view = view;
            _createView = createView;

            _view.CreateBranch += OnShowCreateBranchView;
            _createView.Confirm += OnCreateBranch;
            _createView.Cancel += OnCreateViewCancel;
            _createView.UserInputTextChanged += OnCreateBranchTextChanged;
        }

        ~BranchesPresenter()
        {
            _createView.Close();
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
                                    b => b.Name.Split(new[] {'/'})
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
    }

}
