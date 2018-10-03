using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public interface ISourceCodeHandler : ISourceCodeProvider
    {
        /// <summary>
        /// Replaces the entire module's contents with the specified code.
        /// </summary>
        void SubstituteCode(QualifiedModuleName module, string newCode);
    }

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
        void SetSelection(QualifiedModuleName module, Selection selection);
        CodeString Prettify(QualifiedModuleName module, CodeString original);
        CodeString Prettify(ICodeModule module, CodeString original);
        CodeString GetCurrentLogicalLine(ICodeModule module);
        CodeString GetCurrentLogicalLine(QualifiedModuleName module);
        Selection GetSelection(QualifiedModuleName module);
    }
}
