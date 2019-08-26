using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_Description</c> attribute.
    /// </summary>
    /// <remarks>
    /// This is a class distinct from Member and Variable descriptions, because annotation scoping is annoyingly complicated and Rubberduck has a <strong>much</strong> easier time if module annotations and member annotations don't have the same name.
    /// </remarks>
    [Annotation("ModuleDescription", AnnotationTarget.Module)]
    [FlexibleAttributeValueAnnotation("VB_Description", 1)]
    public sealed class ModuleDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public ModuleDescriptionAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(qualifiedSelection, context, parameters)
        {}
    }
}