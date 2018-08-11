using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IEventSource<out TEventSource>
    {
        TEventSource EventSource { get; }
    }
}
