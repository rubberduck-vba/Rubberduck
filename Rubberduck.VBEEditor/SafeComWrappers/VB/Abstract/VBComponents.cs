using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.Abstract
{
    public abstract class VBComponents
    {
        private static readonly object _locker = new object();
        private static object _components;
        private static IComIds _comIds;

        protected VBComponents(object target, VBType vbType)
        {
            lock (_locker)
            {
                if (_components != null || target == null) { return; }

                _comIds = ComIds.For[vbType];
                AttachEvents(target);
            }
        }

        protected delegate void ComponentAddedDelegate(object vbComponent);
        protected static ComponentAddedDelegate _componentAdded;
        public static event EventHandler<ComponentEventArgs> ComponentAdded;
        protected static void DispatchComponentAdded(object vbComponent)
        {
            Dispatch(ComponentAdded, vbComponent);
        }

        protected delegate void ComponentRemovedDelegate(object vbComponent);
        protected static ComponentRemovedDelegate _componentRemoved;
        public static event EventHandler<ComponentEventArgs> ComponentRemoved;
        protected static void DispatchComponentRemoved(object vbComponent)
        {
            Dispatch(ComponentRemoved, vbComponent);
        }

        protected delegate void ComponentRenamedDelegate(object vbComponent, string oldName);
        protected static ComponentRenamedDelegate _componentRenamed;
        public static event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        protected static void DispatchComponentRenamed(object vbComponent, string oldName)
        {
            Dispatch(ComponentRenamed, vbComponent, oldName);
        }

        protected delegate void ComponentSelectedDelegate(object vbComponent);
        protected static ComponentSelectedDelegate _componentSelected;
        public static event EventHandler<ComponentEventArgs> ComponentSelected;
        protected static void DispatchComponentSelected(object vbComponent)
        {
            Dispatch(ComponentSelected, vbComponent);
        }

        protected delegate void ComponentActivatedDelegate(object vbComponent);
        protected static ComponentActivatedDelegate _componentActivated;
        public static event EventHandler<ComponentEventArgs> ComponentActivated;
        protected static void DispatchComponentActivated(object vbComponent)
        {
            Dispatch(ComponentActivated, vbComponent);
        }

        protected delegate void ComponentReloadedDelegate(object vbComponent);
        protected static ComponentReloadedDelegate _componentReloaded;
        public static event EventHandler<ComponentEventArgs> ComponentReloaded;
        protected static void DispatchComponentReloaded(object vbComponent)
        {
            Dispatch(ComponentReloaded, vbComponent);
        }

        private static void Dispatch(EventHandler<ComponentEventArgs> handler, object vbComponent)
        {
            var localHandler = handler;
            if (localHandler != null)
            {
                var component = VBComponentFactory.Create(vbComponent);
                var project = component.Collection.Parent;
                if (project.Protection != ProjectProtection.Locked)
                {
                    localHandler.Invoke(component, new ComponentEventArgs(project.ProjectId, project, component));
                }
            }
        }

        private static void Dispatch(EventHandler<ComponentRenamedEventArgs> handler, object vbComponent, string oldName)
        {
            var localHandler = handler;
            if (localHandler != null)
            {
                var component = VBComponentFactory.Create(vbComponent);
                var project = component.Collection.Parent;
                if (project.Protection != ProjectProtection.Locked)
                {
                    localHandler.Invoke(component, new ComponentRenamedEventArgs(project.ProjectId, project, component, oldName));
                }
            }
        }

        private static void AttachEvents(object components)
        {
            _components = components;
            _componentAdded = DispatchComponentAdded;
            _componentRemoved = DispatchComponentRemoved;
            _componentRenamed = DispatchComponentRenamed;
            _componentSelected = DispatchComponentSelected;
            _componentActivated = DispatchComponentActivated;
            _componentReloaded = DispatchComponentReloaded;
            ComEventsHelper.Combine(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemAdded, _componentAdded);
            ComEventsHelper.Combine(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemRemoved, _componentRemoved);
            ComEventsHelper.Combine(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemRenamed, _componentRenamed);
            ComEventsHelper.Combine(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemSelected, _componentSelected);
            ComEventsHelper.Combine(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemActivated, _componentActivated);
            ComEventsHelper.Combine(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemReloaded, _componentReloaded);
        }

        public static void DetachEvents()
        {
            lock (_locker)
            {
                if (_components != null)
                {
                    ComEventsHelper.Remove(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemAdded, _componentAdded);
                    ComEventsHelper.Remove(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemRemoved, _componentRemoved);
                    ComEventsHelper.Remove(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemRenamed, _componentRenamed);
                    ComEventsHelper.Remove(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemSelected, _componentSelected);
                    ComEventsHelper.Remove(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemActivated, _componentActivated);
                    ComEventsHelper.Remove(_components, _comIds.VBComponentsEventsGuid, _comIds.ComponentEventDispIds.ItemReloaded, _componentReloaded);
                    _components = null;
                }
            }
        }
    }

    public abstract class VbComponents<T> : VBComponents, ISafeComWrapper<T>, IVBComponents
        where T: class
    {
        private readonly VBComponentsWrapper<T> _comWrapper;
        protected VbComponents(T target, VBType vbType)
            : base(target, vbType)
        {
            _comWrapper = new VBComponentsWrapper<T>(target, Equals, GetHashCode);
        }

        
        public abstract IEnumerator<IVBComponent> GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public abstract int Count { get; }
        public abstract IVBComponent this[object index] { get; }        
        public abstract IVBE VBE { get; }
        public abstract IVBProject Parent { get; }
        public abstract void Remove(IVBComponent item);
        public abstract IVBComponent Add(ComponentType type);
        public abstract IVBComponent Import(string path);
        public abstract IVBComponent AddCustom(string progId);
        public abstract IVBComponent AddMTDesigner(int index = 0);
        public abstract void ImportSourceFile(string path);
        public abstract void RemoveSafely(IVBComponent component);        
        public abstract bool Equals(IVBComponents other);                
        public abstract override int GetHashCode();
        public T Target => _comWrapper.Target;
        public bool IsWrappingNullReference => _comWrapper.IsWrappingNullReference;

        private class VBComponentsWrapper<TItem> : SafeComWrapper<TItem>
            where TItem: class
        {
            private readonly Func<ISafeComWrapper<TItem>, bool> _equals;
            private readonly Func<int> _getHashCode;
            internal VBComponentsWrapper(TItem target, Func<ISafeComWrapper<TItem>, bool> equals, Func<int> getHashCode)
                : base(target)
            {
                _equals = equals;
                _getHashCode = getHashCode;
            }

            public override bool Equals(ISafeComWrapper<TItem> other)
            {
                return _equals.Invoke(other);
            }

            public override int GetHashCode()
            {
                return _getHashCode.Invoke();
            }
        }

        object INullObjectWrapper.Target => _comWrapper.Target;
    }     
}
