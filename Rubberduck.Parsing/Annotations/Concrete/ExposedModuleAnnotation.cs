using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_Exposed</c> attribute.
    /// </summary>
    public sealed class ExposedModuleAnnotation : FixedAttributeValueAnnotationBase
    {
        public ExposedModuleAnnotation()
            : base("Exposed", AnnotationTarget.Module, "VB_Exposed", new[] { "True" })
        {}
    }
}