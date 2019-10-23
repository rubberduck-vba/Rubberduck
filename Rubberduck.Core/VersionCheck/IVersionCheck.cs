using Rubberduck.Settings;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.VersionCheck
{
    public interface IVersionCheck
    {
        Task<Version> GetLatestVersionAsync(GeneralSettings settings, CancellationToken token = default);
        Version CurrentVersion { get; }
    }
}