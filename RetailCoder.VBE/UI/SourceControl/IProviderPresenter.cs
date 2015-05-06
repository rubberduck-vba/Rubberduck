using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface IProviderPresenter
    {
        ISourceControlProvider Provider { get; set; }
    }
}
