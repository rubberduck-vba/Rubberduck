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
        /// <param name="document"><see cref="HostDocument"/> representing the document module</param>
        /// <returns>True if able to get the document, false otherwise</returns>
        /// <inheritdoc cref="GetDocuments"/>
        HostDocument GetDocument(QualifiedModuleName moduleName);

        /// <summary>
        /// Indicate if the host is able to open the document in design view
        /// </summary>
        /// <param name="moduleName">The qualified name of the document module</param>
        /// <returns>True if it can; false otherwise</returns>
        bool CanOpenDocumentDesigner(QualifiedModuleName moduleName);

        /// <summary>
        /// Tries to open a document in design view
        /// </summary>
        /// <param name="moduleName">The qualified name of the document module</param>
        /// <returns>True if the document was opened in design view, false otherwise</returns>
        bool TryOpenDocumentDesigner(QualifiedModuleName moduleName);

        /// <summary>
        /// Get a list of host-specific auto macro identifiers where
        /// applicable. The component type may be used as a bitmask
        /// for where host allows multiple component types. The names
        /// may be left null, indicating that any matches is accepted, but
        /// only one of either may be left null. 
        /// </summary>
        IEnumerable<HostAutoMacro> AutoMacroIdentifiers { get; }
    }
}
