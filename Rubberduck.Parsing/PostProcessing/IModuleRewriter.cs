using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.PostProcessing
{
    public interface IModuleRewriter
    {
        /// <summary>
        /// Rewrites the entire module / applies all changes.
        /// </summary>
        void Rewrite();

        /// <summary>
        /// Removes all tokens for specified <see cref="Declaration"/>. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="target">The <see cref="Declaration"/> to remove.</param>
        /// <remarks>Removes a line that would be left empty by the removal of the declaration.</remarks>
        void Remove(Declaration target);
        /// <summary>
        /// Removes all tokens in specified context. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="target">The <see cref="ParserRuleContext"/> to remove.</param>
        /// <remarks>Removes a line that would be left empty by the removal of the identifier reference token.</remarks>
        void Remove(ParserRuleContext target);
        /// <summary>
        /// Removes all tokens for specified <see cref="IToken"/>. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="target">The <see cref="IToken"/> to remove.</param>
        /// <remarks>Removes a line that would be left empty by the removal of the identifier reference token.</remarks>
        void Remove(IToken target);
        /// <summary>
        /// Removes all tokens for specified <see cref="ITerminalNode"/>. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="target">The <see cref="ITerminalNode"/> to remove.</param>
        /// <remarks>Removes a line that would be left empty by the removal of the identifier reference token.</remarks>
        void Remove(ITerminalNode target);

        /// <summary>
        /// Removes all tokens from the start of the first node to the end of the second node.
        /// </summary>
        /// <param name="start">The start index to remove.</param>
        /// <param name="stop">The end index to remove.</param>
        void RemoveRange(int start, int stop);

        /// <summary>
        /// Replaces all tokens for specified <see cref="Declaration"/> with specified content. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="target">The <see cref="Declaration"/> to replace.</param>
        /// <param name="content">The literal replacement for the declaration.</param>
        /// <remarks>Useful for adding/removing e.g. access modifiers.</remarks>
        void Replace(Declaration target, string content);
        /// <summary>
        /// Replaces all tokens for specified <see cref="ParserRuleContext"/> with specified content. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="target">The <see cref="ParserRuleContext"/> to replace.</param>
        /// <param name="content">The literal replacement for the expression.</param>
        void Replace(ParserRuleContext target, string content);
        /// <summary>
        /// Replaces specified token with specified content. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="token">The <see cref="IToken"/> to replace.</param>
        /// <param name="content">The literal replacement for the expression.</param>
        void Replace(IToken token, string content);
        /// <summary>
        /// Replaces specified token with specified content. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="target">The <see cref="ITerminalNode"/> to replace.</param>
        /// <param name="content">The literal replacement for the expression.</param>
        void Replace(ITerminalNode target, string content);

        /// <summary>
        /// Replaces specified interval with specified content. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="tokenInterval">The <see cref="Interval"/> to replace.</param>
        /// <param name="content">The literal replacement for the expression.</param>
        void Replace(Interval tokenInterval, string content);

        /// <summary>
        /// Inserts specified content before the specified token index in the module. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="tokenIndex">The index of the insertion point in the module's lexer token stream.</param>
        /// <param name="content">The literal content to insert.</param>
        void InsertBefore(int tokenIndex, string content);

        /// <summary>
        /// Inserts specified content after the specified token index in the module. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="tokenIndex">The index of the insertion point in the module's lexer token stream.</param>
        /// <param name="content">The literal content to insert.</param>
        void InsertAfter(int tokenIndex, string content);

        /// <summary>
        /// Gets the text between specified token positions (inclusive).
        /// </summary>
        /// <returns></returns>
        string GetText(int startTokenIndex, int stopTokenIndex);

        /// <summary>
        /// Gets the rewritten module content.
        /// </summary>
        /// <returns></returns>
        string GetText();

        ITokenStream TokenStream { get; }
    }
}