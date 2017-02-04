using System;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.VersionCheck
{
    public interface IVersionCheck
    {
        Task<Version> GetLatestVersionAsync(CancellationToken token = default(CancellationToken));
        Version CurrentVersion { get; }
    }

    public class VersionCheck : IVersionCheck
    {
        private readonly Lazy<Version> _currentVersion;
        public VersionCheck()
        {
            _currentVersion = new Lazy<Version>(() => Assembly.GetExecutingAssembly().GetName().Version);
        }

        private Version _latestVersion;
        public async Task<Version> GetLatestVersionAsync(CancellationToken token = default(CancellationToken))
        {
            if (_latestVersion != default(Version)) { return _latestVersion; }

            using (var client = new HttpClient())
            {
                var url = new Uri("http://rubberduckvba.com/Build/Version/Stable");
                var response = await client.GetAsync(url, token);
                var version = await response.Content.ReadAsStringAsync();
                return _latestVersion = new Version(version);
            }
        }

        public Version CurrentVersion { get { return _currentVersion.Value; } }
    }
}