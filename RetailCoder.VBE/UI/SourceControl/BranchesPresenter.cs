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

        public BranchesPresenter(ISourceControlProvider provider, IBranchesView view)
        {
            _provider = provider;
            _view = view;
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

        private IList<string> GetFriendlyBranchNames(IEnumerable<IBranch> branches)
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
    }

}
