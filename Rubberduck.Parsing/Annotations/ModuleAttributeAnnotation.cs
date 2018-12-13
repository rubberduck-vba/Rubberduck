using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public class ModuleAttributeAnnotation : AttributeAnnotationBase
    {
        public ModuleAttributeAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IReadOnlyList<string> paramaters) 
        :base(AnnotationType.ModuleAttribute, qualifiedSelection, context, paramaters.Skip(1).ToList())
        {
            Attribute = paramaters.FirstOrDefault();
        }

        public override string Attribute { get; }
    }
}