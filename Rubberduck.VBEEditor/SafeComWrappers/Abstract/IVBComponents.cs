using System;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBComponents : ISafeComWrapper, IComCollection<IVBComponent>, IEquatable<IVBComponents>
    {
        //event EventHandler<ComponentEventArgs> ComponentAdded;
        //event EventHandler<ComponentEventArgs> ComponentRemoved;
        //event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        //event EventHandler<ComponentEventArgs> ComponentSelected;
        //event EventHandler<ComponentEventArgs> ComponentActivated;
        //event EventHandler<ComponentEventArgs> ComponentReloaded;

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
    }
}