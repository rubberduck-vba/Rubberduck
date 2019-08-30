using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public class VariableDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {
        public VariableDescriptionAnnotation()
            : base("VariableDescription", AnnotationTarget.Variable, "VB_VarDescription", 1)
        {}
    }
}