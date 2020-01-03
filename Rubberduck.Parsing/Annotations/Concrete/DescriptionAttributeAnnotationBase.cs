using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class DescriptionAttributeAnnotationBase : FlexibleAttributeValueAnnotationBase
    {
        public DescriptionAttributeAnnotationBase(string name, AnnotationTarget target, string attribute, int valueCount)
            : base(name, target, attribute, valueCount)
        { }
    }
}