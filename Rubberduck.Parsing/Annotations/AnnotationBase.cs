using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AnnotationBase : IAnnotation
    {
        public const string ANNOTATION_MARKER = "'@";

        protected AnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection)
        {
            AnnotationType = annotationType;
            QualifiedSelection = qualifiedSelection;
        }

        public AnnotationType AnnotationType { get; }
        public QualifiedSelection QualifiedSelection { get; }

        public virtual bool AllowMultiple { get; } = false;

        public override string ToString() => $"Annotation Type: {AnnotationType}";
    }
}
