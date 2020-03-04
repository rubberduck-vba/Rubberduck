using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings
{
    public interface IRefactoringFactory<TRefactoring>
        where TRefactoring : IRefactoring
    {
        TRefactoring Create();
        void Release(TRefactoring refactoring);
    }

    public interface IRefactoring
    {
        void Refactor();
        void Refactor(QualifiedSelection target);
        void Refactor(Declaration target);
    }
}
