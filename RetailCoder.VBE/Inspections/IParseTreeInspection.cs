using Rubberduck.Parsing;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public interface IParseTreeInspection : IInspection
    {
        ParseTreeResults ParseTreeResults { get; set; }
    }

    public sealed class ParseTreeResults
    {
        public ParseTreeResults()
        {
            ObsoleteCallContexts = Enumerable.Empty<QualifiedContext>();
            ObsoleteLetContexts = Enumerable.Empty<QualifiedContext>();
            ArgListsWithOneByRefParam = Enumerable.Empty<QualifiedContext>();
            EmptyStringLiterals = Enumerable.Empty<QualifiedContext>();
            MalformedAnnotations = Enumerable.Empty<QualifiedContext<VBAParser.AnnotationContext>>();
        }

        public IEnumerable<QualifiedContext> ObsoleteCallContexts;
        public IEnumerable<QualifiedContext> ObsoleteLetContexts;
        public IEnumerable<QualifiedContext> ArgListsWithOneByRefParam;
        public IEnumerable<QualifiedContext> EmptyStringLiterals;
        public IEnumerable<QualifiedContext<VBAParser.AnnotationContext>> MalformedAnnotations;
    }
}
