using System;

namespace Rubberduck.UI
{
    public interface IEnvironmentProvider
    {
        string GetFolderPath(Environment.SpecialFolder folder);
        // ReSharper disable once InconsistentNaming
        OperatingSystem OSVersion { get; }
    }

    //Wrapper to enable unit testing of folder dialogs.
    public class EnvironmentProvider : IEnvironmentProvider
    {
        public string GetFolderPath(Environment.SpecialFolder folder)
        {
            return Environment.GetFolderPath(folder);
        }

        public OperatingSystem OSVersion { get { return Environment.OSVersion; } }
    }
}
