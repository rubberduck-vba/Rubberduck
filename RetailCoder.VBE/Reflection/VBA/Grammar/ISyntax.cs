using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Reflection.VBA.Grammar
{
    internal interface ISyntax
    {
        /// <summary>
        /// Parses an instruction into a syntax node, if possible.
        /// </summary>
        /// <param name="publicScope">The fully-qualified scope of the specified instruction, when the instruction is publicly scoped.</param>.
        /// <param name="localScope">The fully-qualified scope of the specified instruction, when the instruction is locally scoped.</param>
        /// <param name="instruction">A string containing a single instruction.</param>
        /// <returns>
        /// Returns a node representing the specified instruction, 
        /// or <c>null</c> if specified instruction can't be parsed.
        /// </returns>
        SyntaxTreeNode ToNode(string publicScope, string localScope, string instruction);

        bool IsMatch(string publicScope, string localScope, string instruction, out SyntaxTreeNode node);

        /// <summary>
        /// Gets a value indicating whether syntax is specific to a particular parent node.
        /// </summary>
        /// <remarks>
        /// Implementations with this member set to <c>true</c> will not be considered as part of the general grammar.
        /// </remarks>
        bool IsChildNodeSyntax { get; }
    }
}
