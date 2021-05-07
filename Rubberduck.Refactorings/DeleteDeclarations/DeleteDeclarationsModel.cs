using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    //TODO: Revisit Targets property - make it ReadOnlyCollection after DeleteDeclarationsModelEx bites the dust
    public interface IDeleteDeclarationsModel : IRefactoringModel
    {
        void AddDeclarationsToDelete(params Declaration[] targets);

        void AddRangeOfDeclarationsToDelete(IEnumerable<Declaration> targets);

        List<Declaration> Targets { get; }

        bool IndentModifiedModules { set; get; }
    }

    public class DeleteDeclarationsModel : IDeleteDeclarationsModel
    {
        private HashSet<Declaration> _targets = new HashSet<Declaration>();

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

        public List<Declaration> Targets
        {
            set => value.ForEach(v => _targets.Add(v));
            get => new List<Declaration>(_targets);
        }

        /// <summary>
        /// IndentModifiedModules should only be set to 'true' when the DeleteDeclationsRefactoringAction
        /// is the initiating/top-level refactoring action (e.g., RemoveUnusedDeclarationQuickFix)
        /// </summary>
        public bool IndentModifiedModules{ set; get; } = false;
    }
}
