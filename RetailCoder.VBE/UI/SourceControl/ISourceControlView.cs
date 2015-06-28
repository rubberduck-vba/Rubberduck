using System;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlView : IDockableUserControl
    {
        event EventHandler<EventArgs> RefreshData;
        event EventHandler<EventArgs> OpenWorkingDirectory;
        event EventHandler<EventArgs> InitializeNewRepository;
        event EventHandler<EventArgs> DismissMessage;

        string Status { get; set; }
        string FailedActionMessage { get; set; }
        bool FailedActionMessageVisible { get; set; }
    }
}
