using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing.Grammar
{
    public interface IMemberContext : IIdentifierContext
    {
        Attributes Attributes { get; }
        IEnumerable<VBAParser.AnnotationContext> Annotations { get; }

        void Annotate(VBAParser.AnnotationContext annotation);
        void AddAttributes(Attributes attributes);
    }

    public interface IChildContext
    {
        ParserRuleContext ParentContext { get; }
    }
}