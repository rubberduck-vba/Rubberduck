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
        private static readonly bool OldSchool = Environment.OSVersion.Version.Major < 6;

        public IFolderBrowser CreateFolderBrowser(string description)
        {
            return !OldSchool
                ? new ModernFolderBrowser(description) as IFolderBrowser
                : new FolderBrowser(description);
        }

        public IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton)
        {
            return !OldSchool
                ? new ModernFolderBrowser(description, showNewFolderButton) as IFolderBrowser
                : new FolderBrowser(description, showNewFolderButton);
        }

        public IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton, Environment.SpecialFolder rootFolder)
        {
            return !OldSchool
                ? new ModernFolderBrowser(description, showNewFolderButton, rootFolder) as IFolderBrowser
                : new FolderBrowser(description, showNewFolderButton, rootFolder);
        }
    }
}
