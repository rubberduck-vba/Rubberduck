using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteModuleElementsModel : IRefactoringModel
    {
        public DeleteModuleElementsModel()
        {
            Targets = new List<Declaration>();
        }

        public DeleteModuleElementsModel(IEnumerable<Declaration> targets)
        {
            Targets = targets.ToList();
        }

        public List<Declaration> Targets { get; }
    }
}
