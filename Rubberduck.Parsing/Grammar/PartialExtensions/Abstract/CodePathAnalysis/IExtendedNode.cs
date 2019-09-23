using System.Collections.Generic;

namespace Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis
{
    /// <summary>
    /// Marks an extended <see cref="IParseTree"/> node.
    /// </summary>
    public interface IExtendedNode
    {
        /// <summary>
        /// <c>true</c> if the node is traversed in any code path.
        /// </summary>
        bool IsReachable { get; set; }
    }
}