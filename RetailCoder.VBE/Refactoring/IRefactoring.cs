using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactoring
{
    public interface IRefactoring
    {
        Declaration AcquireTarget(QualifiedSelection selection);

        void Refactor();
        void Refactor(QualifiedSelection selection);
        void Refactor(Declaration target);
    }
}
