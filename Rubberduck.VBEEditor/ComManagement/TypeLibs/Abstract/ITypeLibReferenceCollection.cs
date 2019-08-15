using System.Collections.Generic;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeLibReferenceCollection
    {
        int Count { get; }
        ITypeLibReference GetItemByIndex(int index);
        IEnumerator<ITypeLibReference> GetEnumerator();
    }
}