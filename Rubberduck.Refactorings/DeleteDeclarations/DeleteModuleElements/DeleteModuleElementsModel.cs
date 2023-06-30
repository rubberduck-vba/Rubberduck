using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteModuleElementsModel : DeleteDeclarationsModel
    {
        public DeleteModuleElementsModel()
            : base(Enumerable.Empty<Declaration>()) { }

        public DeleteModuleElementsModel(IEnumerable<Declaration> targets)
            : base(targets) { }
    }
}
