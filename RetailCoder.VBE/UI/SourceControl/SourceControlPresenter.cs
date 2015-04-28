using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.SourceControl
{
    public class SourceControlPresenter : DockablePresenterBase
    {
        private readonly IChangesPresenter _changesPresenter;
        private readonly IBranchesPresenter _branchesPresenter;
        private readonly ISourceControlView _view;

        public SourceControlPresenter
            (
                VBE vbe, 
                AddIn addin, 
                ISourceControlView view, 
                IChangesPresenter changesPresenter,
                IBranchesPresenter branchesPresenter           
            ) 
            : base(vbe, addin, view)
        {
            _changesPresenter = changesPresenter;
            _branchesPresenter = branchesPresenter;
            _view = view;

            _view.RefreshData += OnRefreshChildren;
        }

        private void OnRefreshChildren(object sender, EventArgs e)
        {
            RefreshChildren();
        }

        public void RefreshChildren()
        {
            _branchesPresenter.RefreshView();
            _changesPresenter.Refresh();
        }
    }
}
