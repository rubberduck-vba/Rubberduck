using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a member's <c>VB_Description</c> attribute.
    /// </summary>
    public sealed class DescriptionAnnotation : AttributeAnnotationBase
    {
        public DescriptionAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> parameters)
            : base(AnnotationType.Description, qualifiedSelection, context, parameters?.Take(1).ToList())
        {
            Description = AttributeValues?.FirstOrDefault();
            if ((Description?.StartsWith("\"") ?? false) && Description.EndsWith("\""))
            {
                // strip surrounding double quotes
                Description = Description.Substring(1, Description.Length - 2);
            }
        }

        public string Description { get; }
        public override string Attribute => "VB_Description";
    }
}