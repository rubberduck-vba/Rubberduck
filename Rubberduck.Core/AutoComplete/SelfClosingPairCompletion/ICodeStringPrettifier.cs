using Rubberduck.Common;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public interface ICodeStringPrettifier
    {
        /// <summary>
        /// Evaluates whether the specified <see cref="CodeString"/> renders as intended in the VBE.
        /// </summary>
        /// <returns>Returns <c>true</c> if the spacing is unchanged, <c>false</c> if the caret position wouldn't be where it's expected to be.</returns>
        bool IsSpacingUnchanged(CodeString code, CodeString original);
    }
}
