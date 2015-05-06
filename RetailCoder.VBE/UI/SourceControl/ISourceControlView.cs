using System;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlView : IDockableUserControl
    {
        event EventHandler<EventArgs> RefreshData;
        event EventHandler<EventArgs> OpenWorkingDirectory;
        event EventHandler<EventArgs> InitializeNewRepository;

        string Status { get; set; }
    }
}
