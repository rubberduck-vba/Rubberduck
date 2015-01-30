using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public interface ISyntax
    {
        /// <summary>
        /// Parses an instruction into a syntax context, if possible.
        /// </summary>
        /// <param name="publicScope">The fully-qualified scope of the specified instruction, when the instruction is publicly scoped.</param>.
        /// <param name="localScope">The fully-qualified scope of the specified instruction, when the instruction is locally scoped.</param>
        /// <param name="instruction">An instruction.</param>
        /// <returns>
        /// Returns a context representing the specified instruction, 
        /// or <c>null</c> if specified instruction can't be parsed.
        /// </returns>
        SyntaxTreeNode Parse(string publicScope, string localScope, Instruction instruction);

        bool IsMatch(string publicScope, string localScope, Instruction instruction, out SyntaxTreeNode node);

        /// <summary>
        /// Gets a value indicating whether syntax is specific to a particular parent context.
        /// </summary>
        /// <remarks>
        /// Implementations with this member set to <c>true</c> will not be considered as part of the general grammar.
        /// </remarks>
        bool IsChildNodeSyntax { get; }

        SyntaxType Type { get; }
    }
}
