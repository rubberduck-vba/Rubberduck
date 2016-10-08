using System;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract
{
    public interface IVBComponents : ISafeComWrapper, IComCollection<IVBComponent>, IEquatable<IVBComponents>
    {
        IVBE VBE { get; }
        IVBProject Parent { get; }
        void Remove(IVBComponent item);
        IVBComponent Add(ComponentType type);
        IVBComponent Import(string path);
        IVBComponent AddCustom(string progId);
        IVBComponent AddMTDesigner(int index = 0);
        void ImportSourceFile(string path);
        void RemoveSafely(IVBComponent component);
    }
}