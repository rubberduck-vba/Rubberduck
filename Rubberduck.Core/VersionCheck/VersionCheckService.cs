using Newtonsoft.Json;
using Rubberduck.Settings;
using System;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

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

        private async Task<Version> GetGitHubMain()
        {
            var url = new Uri("https://github.com/repos/rubberduck-vba/Rubberduck/releases/latest");
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("rubberduck.version-check"));
                using (var response = await client.GetAsync(url))
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var tagName = (string)JsonConvert.DeserializeObject<dynamic>(content).tag_name;

                    // assumes a tag name like "v2.5.3.0"
                    return new Version(tagName.Substring("v".Length));
                }
            }
        }

        private async Task<Version> GetGitHubNext()
        {
            var url = new Uri("https://github.com/repos/rubberduck-vba/Rubberduck/releases");
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("rubberduck.version-check"));
                using (var response = await client.GetAsync(url))
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var tagName = (string)JsonConvert.DeserializeObject<dynamic>(content)[0].tag_name;

                    // assumes a tag name like "Prerelease-v2.5.2.1234"
                    return new Version(tagName.Substring("Prerelease-v".Length));
                }
            }
        }
    }
}