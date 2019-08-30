using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for ignoring specific inspection results from a specified set of inspections.
    /// </summary>
    public sealed class IgnoreAnnotation : AnnotationBase
    {
        public IgnoreAnnotation()
            : base("Ignore", AnnotationTarget.General, true)
        { }
    }
}
