using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    /// <summary>
    /// Used for specifying a module's <c>VB_Description</c> attribute.
    /// </summary>
    public sealed class ModuleDescriptionAnnotation : AttributeAnnotationBase
    {
        public ModuleDescriptionAnnotation(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context, IEnumerable<string> attributeValues)
            : base(AnnotationType.ModuleDescription, qualifiedSelection, context, attributeValues?.Take(1).ToList())
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