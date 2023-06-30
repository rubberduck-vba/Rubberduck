using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteUDTMembersModel : DeleteDeclarationsModel
    {
        public DeleteUDTMembersModel()
            : base(Enumerable.Empty<Declaration>()) { }

        public DeleteUDTMembersModel(IEnumerable<Declaration> targets)
            : base(targets) { }
    }
}
