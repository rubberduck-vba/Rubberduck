using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IHostApplication : IDisposable
    {
        /// <summary>
        /// Gets the name of the application.
        /// </summary>
        /// <remarks>
        /// This is needed only to circumvent the problem that HostApplicationBase-derived types
        /// cannot be used outside assembly boundaries because the type is generic.
        /// </remarks>
        string ApplicationName { get; }
    }
}
