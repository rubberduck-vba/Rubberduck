using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteEnumMembersModel : DeleteDeclarationsModel
    {
        public DeleteEnumMembersModel()
            : base(Enumerable.Empty<Declaration>()) { }

        public DeleteEnumMembersModel(IEnumerable<Declaration> targets)
            :base(targets){}
    }
}
