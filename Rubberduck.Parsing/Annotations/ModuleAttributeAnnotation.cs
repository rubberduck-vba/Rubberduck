using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public class ModuleAttributeAnnotation : AttributeAnnotationBase
    {
        public ModuleAttributeAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IReadOnlyList<string> parameters) 
        :base(AnnotationType.ModuleAttribute, qualifiedSelection, context, parameters?.Skip(1).ToList())
        {
            Attribute = parameters?.FirstOrDefault() ?? string.Empty;
        }

        public override string Attribute { get; }
    }
}