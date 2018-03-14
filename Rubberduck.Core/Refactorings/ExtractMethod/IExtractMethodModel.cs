using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodModel
    {
        void extract(IEnumerable<Declaration> declarations, QualifiedSelection selection, string selectedCode);
        IExtractedMethod Method { get; }
        IEnumerable<Declaration> DeclarationsToMove { get; }
        IEnumerable<ExtractedParameter> Inputs { get; }
        IEnumerable<Declaration> Locals { get; }
        IEnumerable<ExtractedParameter> Outputs { get; }
        string SelectedCode { get; }
        QualifiedSelection Selection { get; }
        Declaration SourceMember { get; }

        IEnumerable<Selection> RowsToRemove { get; }
        Selection PositionForMethodCall { get; }
        Selection PositionForNewMethod { get; }
        string NewMethodCall { get; }
    }
}
