using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class ModuleInitializeAnnotation : AnnotationBase
    {
        public ModuleInitializeAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.ModuleInitialize, qualifiedSelection)
        {
        }
    }
}
