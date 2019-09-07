using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// This annotation allows specifying arbitrary VB_Attribute entries.
    /// </summary>
    public class ModuleAttributeAnnotation : FlexibleAttributeAnnotationBase
    {
        public ModuleAttributeAnnotation() 
        : base("ModuleAttribute", AnnotationTarget.Module)
        {}
    }
}