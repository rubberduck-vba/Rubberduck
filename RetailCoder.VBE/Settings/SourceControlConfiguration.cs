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

        public SourceControlConfiguration()
        {
            Repositories = new List<Repository>();
            UserName = string.Empty;
            EmailAddress = string.Empty;
            DefaultRepositoryLocation = string.Empty;
        }

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
