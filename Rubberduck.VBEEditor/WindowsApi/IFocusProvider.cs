using System;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.WindowsApi
{
    public interface IFocusProvider : IDisposable
    {
        event EventHandler<WindowChangedEventArgs> FocusChange;
    }
}
