using System;

namespace Rubberduck.UI
{
    public interface IFolderBrowserFactory
    {
        IFolderBrowser CreateFolderBrowser(string description);

        IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton);

        IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton, 
            Environment.SpecialFolder rootFolder);
    }

    public class DialogFactory : IFolderBrowserFactory
    {
        public IFolderBrowser CreateFolderBrowser(string description)
        {
            return new FolderBrowser(description);
        }

        public IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton)
        {
            return new FolderBrowser(description, showNewFolderButton);
        }

        public IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton, Environment.SpecialFolder rootFolder)
        {
            return new FolderBrowser(description, showNewFolderButton, rootFolder);
        }
    }
}
