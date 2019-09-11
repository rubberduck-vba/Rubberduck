using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying the Code Explorer folder a appears under.
    /// </summary>
    public sealed class FolderAnnotation : AnnotationBase
    {
        public FolderAnnotation()
            : base("Folder", AnnotationTarget.Module)
        { }
    }
}
