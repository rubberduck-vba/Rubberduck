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
        public SourceControlPresenter(VBE vbe, AddIn addin, IDockableUserControl control) : base(vbe, addin, control)
        {
        }
    }
}
