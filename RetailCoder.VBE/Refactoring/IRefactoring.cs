using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactoring
{
    public interface IRefactoring
    {
        Declaration AcquireTarget(QualifiedSelection selection);

        void Refactor();
        //void Refactor(QualifiedSelection selection);
        //void Refactor(Declaration target);
    }
}
