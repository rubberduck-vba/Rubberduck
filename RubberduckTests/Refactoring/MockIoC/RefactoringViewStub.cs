using Rubberduck.Refactorings;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class RefactoringViewStub<TModel> : IRefactoringView<TModel>
        where TModel : class
    {
        public object DataContext { get; set; }
        public TModel Model { get; set; }
    }
}
