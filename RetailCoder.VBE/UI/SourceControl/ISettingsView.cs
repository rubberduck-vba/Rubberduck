using System;
using Rubberduck.Settings;

namespace Rubberduck.UI.SourceControl
{
    //string UserName { get; set; }
    //string EmailAddress { get; set; }
    //string DefaultRepositoryLocation { get; set; }

    public interface ISettingsView : ISourceControlUserSettings
    {
        event EventHandler<EventArgs> BrowseDefaultRepositoryLocation; 
        event EventHandler<EventArgs> Save;
        event EventHandler<EventArgs> Cancel; 
        event EventHandler<EventArgs> EditIgnoreFile;
        event EventHandler<EventArgs> EditAttributesFile;
    }
}
