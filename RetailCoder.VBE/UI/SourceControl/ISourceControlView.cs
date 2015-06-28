using System;
using System.Security;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlView : IDockableUserControl
    {
        event EventHandler<EventArgs> RefreshData;
        event EventHandler<EventArgs> OpenWorkingDirectory;
        event EventHandler<EventArgs> InitializeNewRepository;

        string Status { get; set; }

        ISecondarySourceControlPanel SecondaryPanel { get; set; }
        bool SecondaryPanelVisible { get; set; }
    }

    public interface ISecondarySourceControlPanel
    {
        event EventHandler<EventArgs> DismissSecondaryPanel;
    }

    public interface IFailedMessageView : ISecondarySourceControlPanel
    {
        string Message { get; set; }
    }

    public interface ILoginView : ICredentials<SecureString>, ISecondarySourceControlPanel
    {
        event EventHandler Confirm;
        event EventHandler Cancel;
    }
}
