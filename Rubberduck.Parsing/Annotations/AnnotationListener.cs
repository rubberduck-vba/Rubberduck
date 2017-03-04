using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Annotations
{
    public sealed class AnnotationListener : VBAParserBaseListener
    {
        private readonly List<IAnnotation> _annotations;
        private readonly IAnnotationFactory _factory;
        private readonly QualifiedModuleName _qualifiedName;

        public AnnotationListener(IAnnotationFactory factory, QualifiedModuleName qualifiedName)
        {
            _annotations = new List<IAnnotation>();
            _factory = factory;
            _qualifiedName = qualifiedName;
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
            var newAnnotation = _factory.Create(context, new QualifiedSelection(_qualifiedName, context.GetSelection()));
            // It might be an annotation we don't support or a typo.
            if (newAnnotation != null)
            {
                _annotations.Add(newAnnotation);
            }
        }
    }
}
