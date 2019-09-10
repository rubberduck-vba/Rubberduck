using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class UdtMemberContext : IIdentifierContext, IAnnotatedContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out var tokenInterval);
                    return tokenInterval;
                }
            }

            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

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
