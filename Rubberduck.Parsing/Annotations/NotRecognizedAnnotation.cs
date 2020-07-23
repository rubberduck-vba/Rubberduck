namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// The annotation type Rubberduck uses for comments that correctly parse as annotations, but weren't recognized as such.
    /// Since this is not actually an annotation, it has no valid target.
    /// </summary>
    public sealed class NotRecognizedAnnotation : AnnotationBase
    {
        public NotRecognizedAnnotation()
            : base("NotRecognized", 0)
        {}
    }
}