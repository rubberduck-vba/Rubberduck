using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_Exposed</c> attribute.
    /// </summary>
    public sealed class ExposedModuleAnnotation : AnnotationBase, IAttributeAnnotation
    {
        public ExposedModuleAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.Exposed, qualifiedSelection, context)
        {
            
        }

        public string Attribute => "VB_Exposed = True";
    }
}