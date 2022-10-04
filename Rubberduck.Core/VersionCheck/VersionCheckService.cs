using Rubberduck.Core.WebApi;
using Rubberduck.Core.WebApi.Model;
using Rubberduck.Settings;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;

namespace Rubberduck.VersionCheck
{
    public class VersionCheckService : IVersionCheckService
    {
        private readonly IPublicApiClient _api;

        /// <param name="version">That would be the version of the assembly for the <c>_Extension</c> class.</param>
        public VersionCheckService(IPublicApiClient api, Version version)
        {
            _api = api;

            CurrentVersion = version;
#if DEBUG
            IsDebugBuild = true;
#endif
            VersionString = IsDebugBuild
                ? $"{version.Major}.{version.Minor}.{version.Build}.x (debug)"
                : version.ToString();
        }

        private Tag _latestTag;

        public async Task<Version> GetLatestVersionAsync(GeneralSettings settings, CancellationToken token = default)
        {
            if (_latestTag != default) 
            { 
                return _latestTag.Version; 
            }

            try
            {
                var latestTags = await _api.GetLatestTagsAsync();

                _latestTag = latestTags
                    .Where(tag => tag != null 
                        && (!tag.IsPreRelease || settings.IncludePreRelease))
                    .OrderByDescending(tag => tag.Version)
                    .FirstOrDefault();

                return _latestTag?.Version;
            }
            catch
            {
                return CurrentVersion;
            }
        }

        public Version CurrentVersion { get; }
        public bool IsDebugBuild { get; }
        public string VersionString { get; }
    }
}