using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.WindowsApi
{
    public interface IWindowEventProvider : IDisposable
    {
        event EventHandler<WindowChangedEventArgs> FocusChange;
    }
}
