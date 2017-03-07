namespace Rubberduck.UI
{
    public interface IFolderBrowserFactory
    {
        IFolderBrowser CreateFolderBrowser(string description);

        IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton);

        IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton, 
            string rootFolder);
    }

    public class DialogFactory : IFolderBrowserFactory
    {
        private readonly IEnvironmentProvider _environment;
        private readonly bool _oldSchool;

        public DialogFactory(IEnvironmentProvider environment)
        {
            _environment = environment;
            try
            {
                _oldSchool = _environment.OSVersion.Version.Major < 6;
            }
            catch
            {
                // ignored - fall back to "safe" dialog version.
            }
        }

        public IFolderBrowser CreateFolderBrowser(string description)
        {
            return !_oldSchool
                ? new ModernFolderBrowser(_environment, description) as IFolderBrowser
                : new FolderBrowser(_environment, description);
        }

        public IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton)
        {
            return !_oldSchool
                ? new ModernFolderBrowser(_environment, description, showNewFolderButton) as IFolderBrowser
                : new FolderBrowser(_environment, description, showNewFolderButton);
        }

        public IFolderBrowser CreateFolderBrowser(string description, bool showNewFolderButton, string rootFolder)
        {
            return !_oldSchool
                ? new ModernFolderBrowser(_environment, description, showNewFolderButton, rootFolder) as IFolderBrowser
                : new FolderBrowser(_environment, description, showNewFolderButton, rootFolder);
        }
    }
}
