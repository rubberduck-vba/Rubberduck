using Rubberduck.VBEditor;
namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodSelectionValidation
    {
        bool ValidateSelection(QualifiedSelection qualifiedSelection);
        bool ContainsCompilerDirectives { get; set; }
    }
}
