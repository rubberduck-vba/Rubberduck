using Rubberduck.VBEditor;
namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodSelectionValidation
    {
        bool IsSelectionValid(QualifiedSelection qualifiedSelection);
        bool ContainsCompilerDirectives { get; set; }
    }
}
