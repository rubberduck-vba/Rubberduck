using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IDeclarationDeletionGroupsGenerator
    {
        List<IDeclarationDeletionGroup> Generate(IEnumerable<IDeclarationDeletionTarget> declarationDeletionTargets);
    }
}
