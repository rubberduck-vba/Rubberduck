using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteDeclarationsModel : IRefactoringModel
    {
        public DeleteDeclarationsModel(Declaration target)
            : this(new List<Declaration>() { target })
        {}

        public DeleteDeclarationsModel(IEnumerable<Declaration> targets)
        {
            Targets = targets.ToList();
        }

        public IReadOnlyCollection<Declaration> Targets { get; }
        /// <summary>
        /// IndentModifiedModules should only be set to 'true' when the DeleteDeclationsRefactoringAction
        /// is the initiating/top-level refactoring action (e.g., RemoveUnusedDeclarationQuickFix)
        /// </summary>
        public bool IndentModifiedModules{ set; get; } = false;
    }
}
