using System;
using System.Linq;
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

            try
            {
                using (var client = new HttpClient())
                {
                    var url = new Uri("http://rubberduckvba.com/Build/Version/Stable");
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

        public Version CurrentVersion { get { return _currentVersion.Value; } }
    }
}