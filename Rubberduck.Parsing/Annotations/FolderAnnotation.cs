using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying the Code Explorer folder a appears under.
    /// </summary>
    public sealed class FolderAnnotation : AnnotationBase
    {
        public FolderAnnotation(
            QualifiedSelection qualifiedSelection,
            IEnumerable<string> parameters)
            : base(AnnotationType.Folder, qualifiedSelection)
        {
            FolderName = parameters.FirstOrDefault() ?? string.Empty;
        }

        public string FolderName { get; }

        public override string ToString() => $"Folder: {FolderName}";
    }
}
