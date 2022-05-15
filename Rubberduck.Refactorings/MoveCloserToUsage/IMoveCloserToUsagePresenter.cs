using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public interface IMoveCloserToUsagePresenter : IRefactoringPresenter<MoveCloserToUsageModel>
    {
        MoveCloserToUsageModel Show(VariableDeclaration target);
        MoveCloserToUsageModel Model { get; }
    }
}
