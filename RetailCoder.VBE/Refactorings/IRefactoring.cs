using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings
{
    public interface IRefactoring
    {
        void Refactor();
        void Refactor(QualifiedSelection target);
        void Refactor(Declaration target);
    }
}
