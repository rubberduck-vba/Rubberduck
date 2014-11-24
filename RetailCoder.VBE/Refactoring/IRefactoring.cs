using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Refactoring
{
    [ComVisible(false)]
    public interface IRefactoring
    {
        void Refactor(CodeModule module);
    }
}
