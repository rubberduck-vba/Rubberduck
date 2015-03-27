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
        private ChangesPresenter _changesPresenter;
        private BranchesPresenter _branchesPresenter;
        private ISourceControlView _view;

        public SourceControlPresenter
            (
                VBE vbe, 
                AddIn addin, 
                ISourceControlView view, 
                ChangesPresenter changesPresenter,
                BranchesPresenter branchesPresenter           
            ) 
            : base(vbe, addin, view)
        {
            _changesPresenter = changesPresenter;
            _branchesPresenter = branchesPresenter;
            _view = view;
        }
    }
}
