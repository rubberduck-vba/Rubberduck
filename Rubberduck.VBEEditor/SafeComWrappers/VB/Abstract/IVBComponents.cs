using System;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IVBComponents : ISafeEventedComWrapper, IComCollection<IVBComponent>, IEquatable<IVBComponents>
    {
        event EventHandler<ComponentEventArgs> ComponentAdded;
        event EventHandler<ComponentEventArgs> ComponentRemoved;
        event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        event EventHandler<ComponentEventArgs> ComponentSelected;
        event EventHandler<ComponentEventArgs> ComponentActivated;
        event EventHandler<ComponentEventArgs> ComponentReloaded;

        new IVBComponent this[object index] { get; }

        IVBE VBE { get; }
        IVBProject Parent { get; }
        void Remove(IVBComponent item);
        IVBComponent Add(ComponentType type);
        IVBComponent Import(string path);
        IVBComponent AddCustom(string progId);
        IVBComponent ImportSourceFile(string path);

        /// <summary>
        /// Safely removes the specified VbComponent from the collection.
        /// </summary>
        /// <remarks>
        /// UserForms, Class modules, and Standard modules are completely removed from the project.
        /// Since Document type components can't be removed through the VBE, all code in its CodeModule are deleted instead.
        /// </remarks>
        void RemoveSafely(IVBComponent component);
    }
}