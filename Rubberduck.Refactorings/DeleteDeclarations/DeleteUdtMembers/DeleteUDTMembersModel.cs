using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteUDTMembersModel : DeleteDeclarationsModel
    {
        public DeleteUDTMembersModel(IEnumerable<Declaration> targets)
            : base(targets) { }
    }
}
