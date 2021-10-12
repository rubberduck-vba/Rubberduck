using Rubberduck.Settings;
using System;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.VersionCheck
{
    public class VersionCheck : IVersionCheck
    {
        /// <param name="version">That would be the version of the assembly for the <c>_Extension</c> class.</param>
        public VersionCheck(Version version)
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
                var url = settings.IncludePreRelease
                    ? new Uri("http://rubberduckvba.com/build/version/prerelease")
                    : new Uri("http://rubberduckvba.com/build/version/stable");

                using (var client = new HttpClient())
                {
                    var response = await client.GetAsync(url, token);
                    var content = await response.Content.ReadAsStringAsync();
                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(content);
                    var version = doc.DocumentNode.Descendants("body").Single().InnerText.Trim();
                    return _latestVersion = new Version(version);
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