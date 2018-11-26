using System;
using System.Collections.Generic;

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

        /// <summary>
        /// Gets data for the host-specific documents not otherwise exposed via VBIDE API
        /// </summary>
        /// <remarks>
        /// Not all properties are available via VBIDE API. While some properties may be
        /// accessed via the <see cref="IVBComponent.Properties"/>, there are problems
        /// with using those properties when the document is not in a design mode. For
        /// that reason, it's better to get the data using host's object model instead.
        /// </remarks>
        IEnumerable<HostDocument> GetDocuments();

        /// <summary>
        /// Gets data for a host-specific document not otherwise exposed via VBIDE API
        /// </summary>
        /// <param name="moduleName"><see cref="QualifiedModuleName"/> representing a VBComponent object</param>
        /// <returns><see cref="HostDocument"/> data</returns>
        /// <inheritdoc cref="GetDocuments"/>
        HostDocument GetDocument(QualifiedModuleName moduleName);
    }
}
