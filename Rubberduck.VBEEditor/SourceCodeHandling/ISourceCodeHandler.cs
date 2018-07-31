namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public interface ISourceCodeHandler : ISourceCodeProvider
    {
        void SubstituteCode(QualifiedModuleName module, string newCode);
    }
}
