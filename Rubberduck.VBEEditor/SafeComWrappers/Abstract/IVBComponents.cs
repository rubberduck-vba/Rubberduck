using System;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBComponents : ISafeComWrapper, IComCollection<IVBComponent>, IEquatable<IVBComponents>
    {
        new IVBComponent this[object index] { get; }

        IVBE VBE { get; }
        IVBProject Parent { get; }
        void Remove(IVBComponent item);
        IVBComponent Add(ComponentType type);
        IVBComponent Import(string path);
        IVBComponent AddCustom(string progId);
        IVBComponent AddMTDesigner(int index = 0);
        void ImportSourceFile(string path);
        void RemoveSafely(IVBComponent component);

        IVBComponentsEventsSink Events { get; }
        IConnectionPoint ConnectionPoint { get; }
    }
}