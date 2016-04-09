using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class ModuleCleanupAnnotation : AnnotationBase
    {
        public ModuleCleanupAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.TestMethod, qualifiedSelection)
        {
        }
    }
}
