using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// This annotation allows specifying arbitrary VB_Attribute entries.
    /// </summary>
    // FIXME Consider whether the type-hierarchy alone is sufficient to mark this as an Attribute-Annotation
    [Annotation("ModuleAttribute", AnnotationTarget.Module)]
    public class ModuleAttributeAnnotation : FlexibleAttributeAnnotationBase
    {
        public ModuleAttributeAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IReadOnlyList<string> parameters) 
        :base(qualifiedSelection, context, parameters)
        {}
    }
}