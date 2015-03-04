using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.SourceControl
{
    public class GitHubSourceControlDockablePresenter : DockablePresenterBase
    {
        private SourceControlPanel Control { get { return UserControl as SourceControlPanel; } }

        public GitHubSourceControlDockablePresenter(VBE vbe, AddIn addin, SourceControlPanel window) 
            : base(vbe, addin, window)
        {
            
        }
    }
}
