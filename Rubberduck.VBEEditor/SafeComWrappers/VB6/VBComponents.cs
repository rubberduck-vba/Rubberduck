using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBComponents : SafeComWrapper<VB.VBComponents>, IVBComponents
    {
        private static readonly Guid VBComponentsEventsGuid = new Guid("0002E193-0000-0000-C000-000000000046");

        private enum ComponentEventDispId
        {
            ItemAdded = 1,
            ItemRemoved = 2,
            ItemRenamed = 3,
            ItemSelected = 4,
            ItemActivated = 5,
            ItemReloaded = 6
        }

        public VBComponents(VB.VBComponents target, bool rewrapping = false)
            : base(target, rewrapping)
        {
            AttachEvents();
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBProject Parent => new VBProject(IsWrappingNullReference ? null : Target.Parent);

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IVBComponent this[object index] => new VBComponent(IsWrappingNullReference ? null : Target.Item(index));

        public void Remove(IVBComponent item)
        {
            if (item?.Target != null && !IsWrappingNullReference)
            {
                Target.Remove((VB.VBComponent)item.Target);
            }
        }

        public IVBComponent Add(ComponentType type)
        {
            return new VBComponent(IsWrappingNullReference ? null : Target.Add((VB.vbext_ComponentType)type));
        }

        public IVBComponent Import(string path)
        {
            throw new NotImplementedException();
        }

        public IVBComponent AddCustom(string progId)
        {
            return new VBComponent(IsWrappingNullReference ? null : Target.AddCustom(progId));
        }

        public IVBComponent AddMTDesigner(int index = 0)
        {
            throw new NotImplementedException();
        }

        IEnumerator<IVBComponent> IEnumerable<IVBComponent>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IVBComponent>(Target, comObject => new VBComponent((VB.VBComponent)comObject));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator)new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IVBComponent>)this).GetEnumerator();
        }

        public override bool Equals(ISafeComWrapper<VB.VBComponents> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IVBComponents other)
        {
            return Equals(other as SafeComWrapper<VB.VBComponents>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Target);
        }

        public void ImportSourceFile(string path)
        {
            var ext = Path.GetExtension(path);
            var name = Path.GetFileNameWithoutExtension(path);
            if (!File.Exists(path))
            {
                return;
            }

            var codeString = File.ReadAllText(path);
            var codeLines = codeString.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (ext == ComponentTypeExtensions.DocClassExtension)
            {
                var component = this[name];
                if (component.IsWrappingNullReference)
                {
                    throw new IndexOutOfRangeException($"Could not find document component named '{name}'.");
                }
                using (var codeModule = component.CodeModule)
                {
                    codeModule.Clear();
                    codeModule.AddFromString(codeString);
                }
            }
            else if (ext == ComponentTypeExtensions.FormExtension)
            {
                var component = this[name];
                if (component.IsWrappingNullReference)
                {
                    component = Add(ComponentType.UserForm);
                    component.Properties["Caption"].Value = name;
                    component.Name = name;
                }

                var nonAttributeLines = codeLines.TakeWhile(line => !line.StartsWith("Attribute")).Count();
                var attributeLines = codeLines.Skip(nonAttributeLines).TakeWhile(line => line.StartsWith("Attribute")).Count();
                var declarationsStartLine = nonAttributeLines + attributeLines + 1;
                var correctCodeString = string.Join(Environment.NewLine, codeLines.Skip(declarationsStartLine - 1).ToArray());

                using (var codeModule = component.CodeModule)
                {
                    codeModule.Clear();
                    codeModule.AddFromString(correctCodeString);
                }
            }
            else if (ext != ComponentTypeExtensions.FormBinaryExtension)
            {
                Import(path);
            }
        }

        /// <summary>
        /// Safely removes the specified VbComponent from the collection.
        /// </summary>
        /// <remarks>
        /// UserForms, Class modules, and Standard modules are completely removed from the project.
        /// Since Document type components can't be removed through the VBE, all code in its CodeModule are deleted instead.
        /// </remarks>
        public void RemoveSafely(IVBComponent component)
        {
            switch (component.Type)
            {
                case ComponentType.ClassModule:
                case ComponentType.StandardModule:
                case ComponentType.UserForm:
                    Remove(component);
                    break;
                case ComponentType.ActiveXDesigner:
                case ComponentType.Document:
                    using (var codeModule = component.CodeModule)
                    {
                        codeModule.Clear();
                    }
                    break;
                default:
                    break;
            }
        }

        #region Events

        private bool _eventsAttached;
        private void AttachEvents()
        {
            throw new NotImplementedException("Correct the Guid (see comment above), verify the DispIds, then remove this throw.");
            if (!_eventsAttached && !IsWrappingNullReference)
            {
                _componentAdded = OnComponentAdded;
                _componentRemoved = OnComponentRemoved;
                _componentRenamed = OnComponentRenamed;
                _componentSelected = OnComponentSelected;
                _componentActivated = OnComponentActivated;
                _componentReloaded = OnComponentReloaded;
                ComEventsHelper.Combine(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemAdded, _componentAdded);
                ComEventsHelper.Combine(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRemoved, _componentRemoved);
                ComEventsHelper.Combine(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRenamed, _componentRenamed);
                ComEventsHelper.Combine(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemSelected, _componentSelected);
                ComEventsHelper.Combine(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemActivated, _componentActivated);
                ComEventsHelper.Combine(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemReloaded, _componentReloaded);
                _eventsAttached = true;
            }
        }

        private void DetatchEvents()
        {
            if (!_eventsAttached && !IsWrappingNullReference)
            {
                ComEventsHelper.Remove(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemAdded, _componentAdded);
                ComEventsHelper.Remove(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRemoved, _componentRemoved);
                ComEventsHelper.Remove(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRenamed, _componentRenamed);
                ComEventsHelper.Remove(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemSelected, _componentSelected);
                ComEventsHelper.Remove(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemActivated, _componentActivated);
                ComEventsHelper.Remove(Target, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemReloaded, _componentReloaded);
                _eventsAttached = false;
            }
        }

        private delegate void ItemAddedDelegate(VB.VBComponent vbComponent);
        private ItemAddedDelegate _componentAdded;
        public event EventHandler<ComponentEventArgs> ComponentAdded;
        private void OnComponentAdded(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentAdded, vbComponent);
        }

        private delegate void ItemRemovedDelegate(VB.VBComponent vbComponent);
        private ItemRemovedDelegate _componentRemoved;
        public event EventHandler<ComponentEventArgs> ComponentRemoved;
        private void OnComponentRemoved(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentRemoved, vbComponent);
        }

        private delegate void ItemRenamedDelegate(VB.VBComponent vbComponent, string oldName);
        private ItemRenamedDelegate _componentRenamed;
        public event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        private void OnComponentRenamed(VB.VBComponent vbComponent, string oldName)
        {
            var component = new VBComponent(vbComponent);
            var handler = ComponentRenamed;
            if (handler == null)
            {
                component.Dispose();
                return;
            }
            handler.Invoke(this, new ComponentRenamedEventArgs(Parent.ProjectId, Parent, component, oldName));
        }

        private delegate void ItemSelectedDelegate(VB.VBComponent vbComponent);
        private ItemSelectedDelegate _componentSelected;
        public event EventHandler<ComponentEventArgs> ComponentSelected;
        private void OnComponentSelected(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentSelected, vbComponent);
        }

        private delegate void ItemActivatedDelegate(VB.VBComponent vbComponent);
        private ItemActivatedDelegate _componentActivated;
        public event EventHandler<ComponentEventArgs> ComponentActivated;
        private void OnComponentActivated(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentActivated, vbComponent);
        }

        private delegate void ItemReloadedDelegate(VB.VBComponent vbComponent);
        private ItemReloadedDelegate _componentReloaded;
        public event EventHandler<ComponentEventArgs> ComponentReloaded;
        private void OnComponentReloaded(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentReloaded, vbComponent);
        }

        private void OnDispatch(EventHandler<ComponentEventArgs> dispatched, VB.VBComponent component)
        {
            var vbComponent = new VBComponent(component);
            var handler = dispatched;
            if (handler == null)
            {
                vbComponent.Dispose();
                return;
            }
            handler?.Invoke(this, new ComponentEventArgs(Parent.ProjectId, Parent, vbComponent));
        }

        #endregion
    }
}