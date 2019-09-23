using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar.Abstract;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class IdentifierStatementLabelContext : IIdentifierContext, IAnnotatedContext, ILabelNode
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out var tokenInterval);
                    return tokenInterval;
                }
            }

            public bool IsReachable { get; set; }

            public Attributes Attributes { get; } = new Attributes();
            public int AttributeTokenIndex => Start.TokenIndex - 1;

            private readonly List<AnnotationContext> _annotations = new List<AnnotationContext>();
            public IEnumerable<AnnotationContext> Annotations => _annotations;

            public void Annotate(AnnotationContext annotation) => _annotations.Add(annotation);

            public void AddAttributes(Attributes attributes)
            {
                foreach (var node in attributes.Select(attribute => new AttributeNode(attribute.Name, attribute.Values)))
                {
                    Attributes.Add(node);
                }
            }
        }
    }
}
