namespace Rubberduck.VBEditor.SourceCodeHandling
{
    public interface ISourceCodeProvider
    {
        string SourceCode(QualifiedModuleName module);
    }
}
