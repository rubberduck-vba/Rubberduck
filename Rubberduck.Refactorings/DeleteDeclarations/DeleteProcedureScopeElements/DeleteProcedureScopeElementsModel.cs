using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteProcedureScopeElementsModel : DeleteDeclarationsModel
    {
        public DeleteProcedureScopeElementsModel()
            : base(Enumerable.Empty<Declaration>()) { }

        public DeleteProcedureScopeElementsModel(IEnumerable<Declaration> targets)
            : base(targets) { }
    }
}
