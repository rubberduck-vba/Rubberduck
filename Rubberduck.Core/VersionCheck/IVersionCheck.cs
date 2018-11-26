using System;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.VersionCheck
{
    public interface IVersionCheck
    {
        Task<Version> GetLatestVersionAsync(CancellationToken token = default);
        Version CurrentVersion { get; }
    }
}