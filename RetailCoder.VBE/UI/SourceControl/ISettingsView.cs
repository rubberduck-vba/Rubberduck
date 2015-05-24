using System;

namespace Rubberduck.UI.SourceControl
{
    public interface ISettingsView
    {
        string UserName { get; set; }
        string EmailAddress { get; set; }
        string DefaultRepositoryLocation { get; set; }

        event EventHandler<EventArgs> SelectDefaultRepositoryLocation; 
        event EventHandler<EventArgs> Save;
        event EventHandler<EventArgs> Cancel; 
        event EventHandler<EventArgs> EditIgnoreFile;
        event EventHandler<EventArgs> EditAttributesFile;
    }
}
