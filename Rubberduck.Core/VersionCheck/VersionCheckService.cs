using Newtonsoft.Json;
using NLog;
using Rubberduck.Settings;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.VersionCheck
{
    public class VersionCheckService : IVersionCheckService
    {
        private static readonly ILogger _logger = LogManager.GetCurrentClassLogger();

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
            if (_latestVersion != default) 
            {
                return _latestVersion; 
            }

            using (var client = new PublicApiClient())
            {
                var tags = await client.GetLatestTagsAsync(token);

                var next = tags.Single(e => e.IsPreRelease);
                var main = tags.Single(e => !e.IsPreRelease);
                _logger.Info($"Main: v{main.Version.ToString(3)}; Next: v{next.Version.ToString(4)}");

                _latestVersion = settings.IncludePreRelease ? next.Version : main.Version;
                _logger.Info($"Check prerelease: {settings.IncludePreRelease}; latest: v{_latestVersion.ToString(4)}");

                return _latestVersion;
            }
        }

        public Version CurrentVersion { get; }
        public bool IsDebugBuild { get; }
        public string VersionString { get; }
    }
}