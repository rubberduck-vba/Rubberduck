using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class TestMethodAnnotation : AnnotationBase
    {
        public TestMethodAnnotation(VBAParser.AnnotationContext context, AnnotationTargetType targetType, IEnumerable<string> parameters)
            : base(context, AnnotationType.TestMethod, targetType)
        {
        }
    }
}
