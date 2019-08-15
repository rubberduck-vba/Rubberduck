using System.Collections.Generic;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeInfoFunctionCollection
    {
        int Count { get; }
        ITypeInfoFunction GetItemByIndex(int index);
        ITypeInfoFunction Find(string name, PROCKIND procKind);
        IEnumerator<ITypeInfoFunction> GetEnumerator();
    }
}