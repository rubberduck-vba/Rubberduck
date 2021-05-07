using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteUDTMembersModel : IRefactoringModel
    {
        public DeleteUDTMembersModel()
        {
            Targets = new List<Declaration>();
        }

        public DeleteUDTMembersModel(IEnumerable<Declaration> targets)
        {
            Targets = targets.ToList();
        }

        public List<Declaration> Targets { get; }
    }
}
