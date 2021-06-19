using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public interface IDeleteDeclarationsModel : IRefactoringModel
    {
        List<Declaration> Targets { get; }
        bool InsertValidationTODOForRetainedComments { get; }

        string RemoveAllExceptionMessage { get; }

        void SetGroups(List<DeletionGroup> deletionGroups, List<IDeclarationDeletionTarget> deletionTargets);

        List<DeletionGroup> DeletionGroups { get; }

        List<IDeclarationDeletionTarget> DeletionTargets { get; }
    }

    public class DeleteDeclarationsModel : IDeleteDeclarationsModel, IDeleteDeclarationModifyEndOfStatementContentModel
    {
        private HashSet<Declaration> _targets = new HashSet<Declaration>();

        private List<IDeclarationDeletionTarget> _deletionTargets = new List<IDeclarationDeletionTarget>();

        private List<DeletionGroup> _deletionGroups = new List<DeletionGroup>();

        public DeleteDeclarationsModel()
        {}

        public DeleteDeclarationsModel(params Declaration[] targets)
        {
            AddRangeOfDeclarationsToDelete(targets);
        }

        public DeleteDeclarationsModel(IEnumerable<Declaration> targets)
        {
            AddRangeOfDeclarationsToDelete(targets);
        }

        public void AddDeclarationsToDelete(params Declaration[] targets)
        {
            AddRangeOfDeclarationsToDelete(targets);
        }

        public void AddRangeOfDeclarationsToDelete(IEnumerable<Declaration> targets)
        {
            foreach (var t in targets)
            {
                _targets.Add(t);
            }
        }

        public List<Declaration> Targets => new List<Declaration>(_targets);

        /// <summary>
        /// IndentModifiedModules should only be set to 'true' when the DeleteDeclationsRefactoringAction
        /// is the initiating/top-level refactoring action (e.g., RemoveUnusedDeclarationQuickFix)
        /// </summary>
        public bool IndentModifiedModules { set; get; } = false;

        public bool InsertValidationTODOForRetainedComments { set; get; } = true;

        public string RemoveAllExceptionMessage { set;  get; }

        public void SetGroups(List<DeletionGroup> deletionGroups, List<IDeclarationDeletionTarget> deletionTargets)
        {
            _deletionGroups = deletionGroups;
            _deletionTargets = deletionTargets;
        }

        public List<DeletionGroup> DeletionGroups => _deletionGroups;

        public List<IDeclarationDeletionTarget> DeletionTargets => _deletionTargets;

        public bool RemoveDeclarationLogicalLineComment { get; set; } = true;
    }
}
