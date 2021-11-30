using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.NonDisposalDecorators
{
    public class VBComponentsNonDisposalDecorator<T> : NonDisposalDecoratorBase<T>, IVBComponents
        where T : IVBComponents
    {
        public VBComponentsNonDisposalDecorator(T components)
            : base(components)
        { }

        public void AttachEvents()
        {
            WrappedItem.AttachEvents();
        }

        public void DetachEvents()
        {
            WrappedItem.DetachEvents();
        }

        public IEnumerator<IVBComponent> GetEnumerator()
        {
            return WrappedItem.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable) WrappedItem).GetEnumerator();
        }

        public int Count => WrappedItem.Count;

        public event EventHandler<ComponentEventArgs> ComponentAdded
        {
            add => WrappedItem.ComponentAdded += value;
            remove => WrappedItem.ComponentAdded -= value;
        }

        public event EventHandler<ComponentEventArgs> ComponentRemoved
        {
            add => WrappedItem.ComponentRemoved += value;
            remove => WrappedItem.ComponentRemoved -= value;
        }

        public event EventHandler<ComponentRenamedEventArgs> ComponentRenamed
        {
            add => WrappedItem.ComponentRenamed += value;
            remove => WrappedItem.ComponentRenamed -= value;
        }

        public event EventHandler<ComponentEventArgs> ComponentSelected
        {
            add => WrappedItem.ComponentSelected += value;
            remove => WrappedItem.ComponentSelected -= value;
        }

        public event EventHandler<ComponentEventArgs> ComponentActivated
        {
            add => WrappedItem.ComponentActivated += value;
            remove => WrappedItem.ComponentActivated -= value;
        }

        public event EventHandler<ComponentEventArgs> ComponentReloaded
        {
            add => WrappedItem.ComponentReloaded += value;
            remove => WrappedItem.ComponentReloaded -= value;
        }

        public IVBComponent this[object index] => WrappedItem[index];

        public IVBE VBE => WrappedItem.VBE;

        public IVBProject Parent => WrappedItem.Parent;

        public void Remove(IVBComponent item)
        {
            WrappedItem.Remove(item);
        }

        public IVBComponent Add(ComponentType type)
        {
            return WrappedItem.Add(type);
        }

        public IVBComponent Import(string path)
        {
            return WrappedItem.Import(path);
        }

        public IVBComponent AddCustom(string progId)
        {
            return WrappedItem.AddCustom(progId);
        }

        public IVBComponent ImportSourceFile(string path)
        {
            return WrappedItem.ImportSourceFile(path);
        }

        public void RemoveSafely(IVBComponent component)
        {
            WrappedItem.RemoveSafely(component);
        }

        public bool Equals(IVBComponents other)
        {
            return WrappedItem.Equals(other);
        }
    }
}