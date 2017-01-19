using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        /// <summary>
        /// Parse tree inspections have their results property-injected.
        /// </summary>
        void SetResults(IEnumerable<QualifiedContext> results);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TContext"></typeparam>
    public interface IParseTreeInspection<TContext> : IParseTreeInspection where TContext : ParserRuleContext
    {
        IEnumerable<QualifiedContext<TContext>> ParseTreeResults { get; }
    }
}
