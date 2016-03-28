using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AnnotationBase : IAnnotation
    {
        private readonly VBAParser.AnnotationContext _context;
        private readonly AnnotationType _annotationType;
        private readonly AnnotationTargetType _targetType;

        public const string ANNOTATION_MARKER = "'@";

        public AnnotationBase(VBAParser.AnnotationContext context, AnnotationType annotationType, AnnotationTargetType targetType)
        {
            _context = context;
            _annotationType = annotationType;
            _targetType = targetType;
        }

        public VBAParser.AnnotationContext Context
        {
            get
            {
                return _context;
            }
        }

        public AnnotationType AnnotationType
        {
            get
            {
                return _annotationType;
            }
        }

        public AnnotationTargetType TargetType
        {
            get
            {
                return _targetType;
            }
        }

        public override string ToString()
        {
            return string.Format("Annotation Type: {0}. TargetType: {1}.", _annotationType, _targetType);
        }
    }
}
