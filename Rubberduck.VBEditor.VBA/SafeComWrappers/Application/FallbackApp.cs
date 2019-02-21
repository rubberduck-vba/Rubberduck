using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public sealed class FallbackApp : IHostApplication
    {
        public FallbackApp(IVBE vbe) { }
        public string ApplicationName => "(unknown)";
        public IEnumerable<HostDocument> GetDocuments() => Enumerable.Empty<HostDocument>();
        public HostDocument GetDocument(QualifiedModuleName moduleName)
        {
            return null;
        }
        public bool CanOpenDocumentDesigner(QualifiedModuleName moduleName) => false;
        public bool TryOpenDocumentDesigner(QualifiedModuleName moduleName) => false;
        public IEnumerable<HostAutoMacro> AutoMacroIdentifiers => Enumerable.Empty<HostAutoMacro>();
        public void Dispose() { }
    }
}
