using Rubberduck.Parsing;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections
{
    internal interface IParseTreeInspection : IInspection
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
        }

        public IEnumerable<QualifiedContext> ObsoleteCallContexts;
        public IEnumerable<QualifiedContext> ObsoleteLetContexts;
        public IEnumerable<QualifiedContext> ArgListsWithOneByRefParam;
        public IEnumerable<QualifiedContext> EmptyStringLiterals;
    }
}
