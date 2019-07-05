using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    /// <summary>
    /// An object that can rewrite a module's contents.
    /// </summary>
    public interface ISourceCodeHandler : ISourceCodeProvider
    {
        /// <summary>
        /// Replaces the entire module's contents with the specified code.
        /// </summary>
        void SubstituteCode(QualifiedModuleName module, string newCode);
    }

    /// <summary>
    /// An object that can manipulate the code in a CodePane.
    /// </summary>
    public interface ICodePaneHandler : ISourceCodeHandler
    {
        /// <summary>
        /// Replaces one or more specific line(s) in the specified module.
        /// </summary>
        void SubstituteCode(QualifiedModuleName module, CodeString newCode);
        /// <summary>
        /// Replaces one or more specific line(s) in the specified module.
        /// </summary>
        void SubstituteCode(ICodeModule module, CodeString newCode);
        void SetSelection(ICodeModule module, Selection selection);
        CodeString Prettify(QualifiedModuleName module, CodeString original);
        CodeString Prettify(ICodeModule module, CodeString original);
        CodeString GetCurrentLogicalLine(ICodeModule module);
    }
}
