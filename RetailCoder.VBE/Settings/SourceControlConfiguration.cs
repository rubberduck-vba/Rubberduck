using System.Collections.Generic;
using Rubberduck.SourceControl;

namespace Rubberduck.Settings
{
    public interface ISourceControlUserSettings
    {
        string UserName { get; set; }
        string EmailAddress { get; set; }
        string DefaultRepositoryLocation { get; set; }
    }

    public class SourceControlConfiguration : ISourceControlUserSettings
    {
        public string UserName { get; set; }
        public string EmailAddress { get; set; }
        public string DefaultRepositoryLocation { get; set; }
        public List<Repository> Repositories;

        public SourceControlConfiguration() : this(string.Empty, string.Empty, string.Empty, new List<Repository>()) { }

        public SourceControlConfiguration
            (
                string username, 
                string email, 
                string defaultRepoLocation,
                List<Repository> repositories
            )
        {
            Repositories = repositories;
            UserName = username;
            EmailAddress = email;
            DefaultRepositoryLocation = defaultRepoLocation;
        }
    }
}
