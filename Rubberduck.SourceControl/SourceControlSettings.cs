using System.Collections.Generic;

namespace Rubberduck.SourceControl
{
    public interface ISourceControlSettings
    {
        string UserName { get; set; }
        string EmailAddress { get; set; }
        string DefaultRepositoryLocation { get; set; }
        List<Repository> Repositories { get; set; }
        string CommandPromptLocation { get; set; }
    }

    public class SourceControlSettings : ISourceControlSettings
    {
        public string UserName { get; set; }
        public string EmailAddress { get; set; }
        public string DefaultRepositoryLocation { get; set; }
        public List<Repository> Repositories { get; set; }
        public string CommandPromptLocation { get; set; }

        public SourceControlSettings() : this(string.Empty, string.Empty, string.Empty, new List<Repository>(), "cmd.exe") { }

        public SourceControlSettings
            (
                string username, 
                string email, 
                string defaultRepoLocation,
                List<Repository> repositories,
                string commandPromptLocation
            )
        {
            UserName = username;
            EmailAddress = email;
            DefaultRepositoryLocation = defaultRepoLocation;
            Repositories = repositories;
            CommandPromptLocation = commandPromptLocation;
        }
    }
}
