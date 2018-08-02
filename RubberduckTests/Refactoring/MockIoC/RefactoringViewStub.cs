using Rubberduck.Refactorings;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class RefactoringViewStub<TModel> : IRefactoringView<TModel>
        where TModel : class
    {
        public virtual object DataContext { get; set; }
        public virtual TModel Model { get; set; }
    }
}
