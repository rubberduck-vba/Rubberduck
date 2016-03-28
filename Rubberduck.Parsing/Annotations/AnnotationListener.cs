using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class AnnotationListener : VBAParserBaseListener
    {
        private readonly List<IAnnotation> _annotations;
        private readonly IAnnotationFactory _factory;

        public AnnotationListener(IAnnotationFactory factory)
        {
            _annotations = new List<IAnnotation>();
            _factory = factory;
        }

        public IEnumerable<IAnnotation> Annotations
        {
            get
            {
                return _annotations;
            }
        }

        public override void ExitAnnotation([NotNull] VBAParser.AnnotationContext context)
        {
            var newAnnotation = _factory.Create(context, AnnotationTargetType.Unknown);
            _annotations.Add(newAnnotation);
        }
    }
}
