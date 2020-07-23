using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Annotations.Concrete
{
    public abstract class DescriptionAttributeAnnotationBase : FlexibleAttributeValueAnnotationBase
    {
        public DescriptionAttributeAnnotationBase(string name, AnnotationTarget target, string attribute)
            : base(name, target, attribute, 1, new List<AnnotationArgumentType> { AnnotationArgumentType.Text })
        {}
    }
}