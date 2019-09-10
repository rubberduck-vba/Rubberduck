using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar.PartialExtensions.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        // holds module-scoped annotations
        public partial class ModuleAttributesContext : IAnnotatedContext 
        {
            public Attributes Attributes { get; } = new Attributes();

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public int AttributeTokenIndex => this.Stop.TokenIndex + 1;

            public void Annotate(AnnotationContext annotation) => _annotations.Add(annotation);

            public void AddAttributes(Attributes attributes)
            {
                foreach (var attribute in attributes)
                {
                    Attributes.Add(new AttributeNode(attribute.Name, attribute.Values));
                }
            }
        }
    }
}
