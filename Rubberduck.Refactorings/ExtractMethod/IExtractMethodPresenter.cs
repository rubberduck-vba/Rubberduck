using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodPresenter : IRefactoringPresenter<ExtractMethodModel>
    {
        //ExtractMethodModel Show(IExtractMethodModel methodModel, IExtractMethodProc extractMethodProc);
        ExtractMethodModel Model { get; }
    }
}