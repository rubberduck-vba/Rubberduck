using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AnnotationBase : IAnnotation
    {
        private readonly AnnotationType _annotationType;
        private readonly QualifiedSelection _qualifiedSelection;

        public const string ANNOTATION_MARKER = "'@";

        public AnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection)
        {
            _annotationType = annotationType;
            _qualifiedSelection = qualifiedSelection;
        }

        public AnnotationType AnnotationType
        {
            get
            {
                return _annotationType;
            }
        }

        public QualifiedSelection QualifiedSelection
        {
            get
            {
                return _qualifiedSelection;
            }
        }

        public override string ToString()
        {
            return string.Format("Annotation Type: {0}", _annotationType);
        }
    }
}
