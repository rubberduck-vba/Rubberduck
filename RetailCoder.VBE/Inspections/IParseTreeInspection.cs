using Rubberduck.Parsing;
using System.Collections.Generic;

namespace Rubberduck.Inspections
{
    internal interface IParseTreeInspection : IInspection
    {
        ParseTreeResults ParseTreeResults { get; set; }
    }

    public sealed class ParseTreeResults
    {
        public IList<QualifiedContext> ObsoleteCallContexts;
        public IList<QualifiedContext> ObsoleteLetContexts;
        public IList<QualifiedContext> ArgListsWithOneByRefParam;
        public IList<QualifiedContext> EmptyStringLiterals;
    }
}