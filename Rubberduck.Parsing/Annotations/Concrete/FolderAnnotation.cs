namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying the Code Explorer folder a appears under.
    /// </summary>
    public sealed class FolderAnnotation : AnnotationBase
    {
        public FolderAnnotation()
            : base("Folder", AnnotationTarget.Module, 1, 1)
        {}
    }
}
