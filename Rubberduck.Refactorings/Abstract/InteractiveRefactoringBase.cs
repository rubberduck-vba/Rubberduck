using Rubberduck.VBEditor.Utility;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings
{
    public abstract class InteractiveRefactoringBase<TModel> : RefactoringBase
        where TModel : class, IRefactoringModel
    {
        private readonly IRefactoringUserInteraction<TModel> _userInteraction;

        protected InteractiveRefactoringBase( 
            ISelectionProvider selectionProvider, 
            IRefactoringUserInteraction<TModel> userInteraction) 
        :base(selectionProvider)
        {
            _userInteraction = userInteraction;
        }

        public override void Refactor(Declaration target)
        {
            Refactor(InitializeModel(target));
        }

        protected void Refactor(TModel initialModel)
        {
            var model = _userInteraction.UserModifiedModel(initialModel);
            RefactorImpl(model);
        }

        protected abstract TModel InitializeModel(Declaration target);
        protected abstract void RefactorImpl(TModel model);
    }
}