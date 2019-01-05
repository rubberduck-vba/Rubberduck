using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public class VariableDescriptionAnnotation : DescriptionAttributeAnnotationBase
    {    
        public VariableDescriptionAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.VariableDescription, qualifiedSelection, context, parameters)
        {}
    }
}