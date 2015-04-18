using System;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlView : IDockableUserControl
    {
        event EventHandler<EventArgs> RefreshData;
    }
}
