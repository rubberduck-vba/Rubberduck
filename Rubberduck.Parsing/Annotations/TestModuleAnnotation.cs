using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class TestModuleAnnotation : AnnotationBase
    {
        public TestModuleAnnotation(VBAParser.AnnotationContext context, AnnotationTargetType targetType)
            : base(context, AnnotationType.TestModule, targetType)
        {
        }
    }
}
