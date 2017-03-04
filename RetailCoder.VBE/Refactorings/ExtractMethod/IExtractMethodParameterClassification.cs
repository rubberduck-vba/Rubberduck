using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodParameterClassification
    {
        IEnumerable<ExtractedParameter> ExtractedParameters { get; }
        void classifyDeclarations(QualifiedSelection selection, Declaration item);
        IEnumerable<Declaration> DeclarationsToMove { get; }
        IEnumerable<Declaration> ExtractedDeclarations { get; }
    }
}
