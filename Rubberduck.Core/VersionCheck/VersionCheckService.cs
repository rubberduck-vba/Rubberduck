using Newtonsoft.Json;
using Rubberduck.Settings;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Octokit;
using System.Net;

namespace Rubberduck.VersionCheck
{
    public class VersionCheckService : IVersionCheckService
    {
        /// <param name="version">That would be the version of the assembly for the <c>_Extension</c> class.</param>
        public VersionCheckService(Version version)
        {
            CurrentVersion = version;
#if DEBUG
            IsDebugBuild = true;
#endif
            VersionString = IsDebugBuild
                ? $"{version.Major}.{version.Minor}.{version.Build}.x (debug)"
                : version.ToString();
        }

        private Version _latestVersion;
        public async Task<Version> GetLatestVersionAsync(GeneralSettings settings, CancellationToken token = default)
        {
            if (_latestVersion != default) { return _latestVersion; }

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                return settings.IncludePreRelease 
                    ? await GetGitHubNext() 
                    : await GetGitHubMain();
            }
            catch
            {
                return _latestVersion;
            }
        }

        public Version CurrentVersion { get; }
        public bool IsDebugBuild { get; }
        public string VersionString { get; }

        private const string GitHubOrgName = "rubberduck-vba";
        private const string GitHubRepoName = "Rubberduck";
        private const string UserAgentProductName = "Rubberduck";
        private GitHubClient GetGitHubClient() => new GitHubClient(new ProductHeaderValue(UserAgentProductName, CurrentVersion.ToString(3)));

        private async Task<Version> GetGitHubMain()
        {
            var client = GetGitHubClient();
            var response = await client.Repository.Release.GetLatest(GitHubOrgName, GitHubRepoName);
            var tagName = response.TagName;

            return new Version(tagName.Substring("v".Length));
        }

        private async Task<Version> GetGitHubNext()
        {
            var client = GetGitHubClient();
            var response = await client.Repository.Release.GetAll(GitHubOrgName, GitHubRepoName);
            var tagName = response.FirstOrDefault()?.TagName ?? "Prerelease-v0.0.0";

            return new Version(tagName.Substring("Prerelease-v".Length));
        }
    }
}