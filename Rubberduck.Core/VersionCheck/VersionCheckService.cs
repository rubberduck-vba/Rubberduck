using Newtonsoft.Json;
using Rubberduck.Settings;
using System;
using System.Linq;
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
                using (var client = new PublicApiClient())
                {
                    var tags = await client.GetLatestTagsAsync(token);
                    var next = tags.Single(e => e.IsPreRelease).Name;
                    var main = tags.Single(e => !e.IsPreRelease).Name;

                    var version = settings.IncludePreRelease
                        ? next.Substring("Prerelease-v".Length)
                        : main.Substring("v".Length);

                    _latestVersion = new Version(version);
                    return _latestVersion;
                }
            }
            catch
            {
                return _latestVersion;
            }
        }

        public Version CurrentVersion { get; }
        public bool IsDebugBuild { get; }
        public string VersionString { get; }
    }
}