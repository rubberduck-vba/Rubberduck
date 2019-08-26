using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class DescriptionAttributeAnnotationBase : FlexibleAttributeValueAnnotationBase
    {
        public DescriptionAttributeAnnotationBase(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> attributeValues)
            : base(qualifiedSelection, context, attributeValues?.Take(1).ToList())
        {
            Description = AttributeValues?.FirstOrDefault();
            if ((Description?.StartsWith("\"") ?? false) && Description.EndsWith("\""))
            {
                // strip surrounding double quotes
                Description = Description.Substring(1, Description.Length - 2);
            }
        }

        public string Description { get; }
    }
}