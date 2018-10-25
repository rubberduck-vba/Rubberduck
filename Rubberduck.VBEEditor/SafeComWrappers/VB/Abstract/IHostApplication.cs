using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

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
        IEnumerable<IHostDocument> GetDocuments();

        /// <summary>
        /// Gets data for a host-specific document not otherwise exposed via VBIDE API
        /// </summary>
        /// <param name="moduleName"><see cref="QualifiedModuleName"/> representing a VBComponent object</param>
        /// <returns><see cref="IHostDocument"/> data</returns>
        /// <inheritdoc cref="GetDocuments"/>
        IHostDocument GetDocument(QualifiedModuleName moduleName);
    }

    public enum DocumentState
    {
        /// <summary>
        /// The document is not open and its accompanying <see cref="IVBComponent"/> may not be available.
        /// </summary>
        Closed,
        /// <summary>
        /// The document is open in design mode.
        /// </summary>
        DesignView,
        /// <summary>
        /// The document is open in non-design mode. It may not be safe to parse the document in this state.
        /// </summary>
        ActiveView
    }

    public interface IHostDocument : IDisposable
    {
        string Name { get; }
        string ClassName { get; }
        WeakReference<object> Target { get; }
        DocumentState State { get; }
    }

    public class HostDocument : IHostDocument
    {
        public HostDocument(string name, string className, object target, DocumentState state)
        {
            Name = name;
            ClassName = className;
            Target = new WeakReference<object>(target);
            State = state;
        }

        public string Name { get; }
        public string ClassName { get; }
        public WeakReference<object> Target { get; }
        public DocumentState State { get; }

        private bool _disposed;
        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;

                if (Target.TryGetTarget(out var target))
                {
                    Marshal.ReleaseComObject(target);
                }
            }
        }
    }
}
