using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteModuleElementsModel : DeleteDeclarationsModel
    {
        public DeleteModuleElementsModel(IEnumerable<Declaration> targets)
            : base(targets) { }
    }
}
