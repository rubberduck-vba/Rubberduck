using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteEnumMembersModel : IRefactoringModel
    {
        public DeleteEnumMembersModel()
        {
            Targets = new List<Declaration>();
        }

        public DeleteEnumMembersModel(IEnumerable<Declaration> targets)
        {
            Targets = targets.ToList();
        }

        public List<Declaration> Targets { get; }
    }
}
