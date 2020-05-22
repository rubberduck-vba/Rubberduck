namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for all annotations not recognized by RD.
    /// Since this is not actually an annotation, it has no valid target
    /// </summary>
    public sealed class NotRecognizedAnnotation : AnnotationBase
    {
        public NotRecognizedAnnotation()
            : base("NotRecognized", 0)
        {}
    }
}