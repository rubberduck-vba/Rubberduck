using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public interface ICodeStringPrettifier
    {
        /// <summary>
        /// Forces the VBE to "prettify" the specified line of code.
        /// </summary>
        /// <param name="module">The module to modify.</param>
        /// <param name="original">The line of code being edited, and current caret position.</param>
        /// <returns>The "prettified" code and caret position.</returns>
        CodeString Prettify(ICodeModule module, CodeString original);
    }
}
