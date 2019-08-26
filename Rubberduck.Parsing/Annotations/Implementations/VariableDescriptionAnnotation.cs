using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    [Annotation("VariableDescription", AnnotationTarget.Variable)]
    [FlexibleAttributeValueAnnotation("VB_VarDescription", 1)]
    public class VariableDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {    
        public VariableDescriptionAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(qualifiedSelection, context, parameters)
        {}
    }
}