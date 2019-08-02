using Antlr4.Runtime.Tree;
using System;

namespace Rubberduck.Parsing.Grammar
{
    /// <summary>
    /// Provides access to common Methods properies
    /// </summary>
    public interface IMethodStmtContext
    {
        /// <summary>
        /// The name of method in context
        /// </summary>
        string MethodName { get; }

        /// <summary>
        /// The kind of method in context
        /// </summary>
        string MethodKind { get; }
    }
}