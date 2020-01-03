using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

// ReSharper disable once CheckNamespace
namespace Rubberduck.Parsing.Grammar
{
    public interface IAnnotatedContext
    {
        Attributes Attributes { get; }
        IEnumerable<VBAParser.AnnotationContext> Annotations { get; }

        /// <summary>
        /// The token index any missing attribute would be inserted at.
        /// </summary>
        int AttributeTokenIndex { get; }

        void Annotate(VBAParser.AnnotationContext annotation);
        void AddAttributes(Attributes attributes);
    }

    public interface IAnnotatingContext
    {
        ParserRuleContext AnnotatedContext { get; }
        string AnnotationType { get; }
    }
}