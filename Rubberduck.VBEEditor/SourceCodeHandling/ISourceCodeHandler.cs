namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public interface ISourceCodeHandler : ISourceCodeProvider
    {
        void SubstituteCode(QualifiedModuleName module, string newCode);
    }

    public interface ICodePaneHandler : ISourceCodeHandler
    {
        void SetSelection(QualifiedModuleName module, Selection selection);
        CodeString Prettify(QualifiedModuleName module, CodeString original);
        CodeString GetCurrentLogicalLine(QualifiedModuleName module);
    }
}
