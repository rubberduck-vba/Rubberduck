using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_Exposed</c> attribute.
    /// </summary>
    public sealed class ExposedModuleAnnotation : AnnotationBase, IAttributeAnnotation
    {
        public ExposedModuleAnnotation(QualifiedSelection qualifiedSelection, IEnumerable<string> parameters)
            : base(AnnotationType.Exposed, qualifiedSelection)
        {
            
        }

        public string Attribute => "VB_Exposed = True";
    }
}