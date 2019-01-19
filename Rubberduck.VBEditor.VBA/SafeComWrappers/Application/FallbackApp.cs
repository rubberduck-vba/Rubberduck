using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public sealed class FallbackApp : IHostApplication
    {
        public FallbackApp(IVBE vbe) { }
        public string ApplicationName => "(unknown)";
        public IEnumerable<HostDocument> GetDocuments() => null;
        public bool TryGetDocument(QualifiedModuleName moduleName, out HostDocument document)
        {
            document = null;
            return false;
        }
        public bool CanOpenDocumentDesigner(QualifiedModuleName moduleName) => false;
        public bool TryOpenDocumentDesigner(QualifiedModuleName moduleName) => false;
        public void Dispose() { }
    }
}
