using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class TestInitializeAnnotation : AnnotationBase
    {
        public TestInitializeAnnotation(VBAParser.AnnotationContext context, AnnotationTargetType targetType, IEnumerable<string> parameters)
            : base(context, AnnotationType.TestMethod, targetType)
        {
        }
    }
}
