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

            RefreshView();
        }

        public void RefreshView()
        {
            _view.Branches = _provider.Branches.Select(b => b.FriendlyName).ToList();
            _view.CurrentBranch = _provider.CurrentBranch.FriendlyName;
        }
    }
}
