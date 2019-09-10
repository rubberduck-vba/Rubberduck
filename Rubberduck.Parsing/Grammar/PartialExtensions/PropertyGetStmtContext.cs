using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar.Abstract;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class PropertyGetStmtContext : IIdentifierContext, IAnnotatedContext
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
            public int AttributeTokenIndex => block()?.Start.TokenIndex ?? Stop.TokenIndex + 1;

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
