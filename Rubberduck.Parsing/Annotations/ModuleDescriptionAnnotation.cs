using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_Description</c> attribute.
    /// </summary>
    public sealed class ModuleDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public ModuleDescriptionAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.ModuleDescription, qualifiedSelection, context, parameters)
        {}
    }
}