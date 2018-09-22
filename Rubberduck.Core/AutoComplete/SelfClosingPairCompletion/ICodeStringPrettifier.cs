using Rubberduck.Common;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public interface ICodeStringPrettifier
    {
        /// <summary>
        /// Forces the VBE to "prettify" the specified line of code.
        /// </summary>
        /// <param name="original">The line of code being edited, and current caret position.</param>
        /// <returns>The "prettified" code and caret position.</returns>
        CodeString Prettify(CodeString original);
        /// <summary>
        /// Evaluates whether the specified <see cref="CodeString"/> renders as intended in the VBE.
        /// </summary>
        /// <returns>Returns <c>true</c> if the spacing is unchanged, <c>false</c> if the caret position wouldn't be where it's expected to be.</returns>
        bool IsSpacingUnchanged(CodeString code, CodeString original);
    }
}
