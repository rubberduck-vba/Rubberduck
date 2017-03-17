using Antlr4.Runtime;
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
        /// Inserts specified content at the specified token index in the module. Use <see cref="Rewrite"/> method to apply changes.
        /// </summary>
        /// <param name="content">The literal content to insert.</param>
        /// <param name="tokenIndex">The index of the insertion point in the module's lexer token stream.</param>
        void InsertAtIndex(string content, int tokenIndex);

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
    }
}