using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;
using VBAIA = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    public class VBComponents : SafeComWrapper<VBAIA.VBComponents>, IVBComponents
    {
        private static readonly Guid VBComponentsEventsGuid = new Guid("0002E116-0000-0000-C000-000000000046");
        private static readonly object Locker = new object();
        private static VBAIA.VBComponents _components;

        private enum ComponentEventDispId
        {
            ItemAdded = 1,
            ItemRemoved = 2,
            ItemRenamed = 3,
            ItemSelected = 4,
            ItemActivated = 5,
            ItemReloaded = 6
        }

        public VBComponents(VBAIA.VBComponents target) : base(target)
        {
            if (_components == null)
            {
                AttachEvents(target);
            }
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;
        public IVBProject Parent => new VBProject(IsWrappingNullReference ? null : Target.Parent);
        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);
        public IVBComponent this[object index] => new VBComponent(IsWrappingNullReference ? null : Target.Item(index));

        public void Remove(IVBComponent item)
        {
            if (item?.Target != null && !IsWrappingNullReference && item.Type != ComponentType.Document)
            {
                Target.Remove((VBAIA.VBComponent)item.Target);
            }
        }

        public IVBComponent Add(ComponentType type)
        {
            return new VBComponent(IsWrappingNullReference ? null : Target.Add((VBAIA.vbext_ComponentType)type));
        }

        public IVBComponent Import(string path)
        {
            return new VBComponent(IsWrappingNullReference ? null : Target.Import(path));
        }

        public IVBComponent AddCustom(string progId)
        {
            return new VBComponent(IsWrappingNullReference ? null : Target.AddCustom(progId));
        }

        public IVBComponent AddMTDesigner(int index = 0)
        {
            return new VBComponent(IsWrappingNullReference ? null : Target.AddMTDesigner(index));
        }

        IEnumerator<IVBComponent> IEnumerable<IVBComponent>.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<IVBComponent>(null, o => new VBComponent(null))
                : new ComWrapperEnumerator<IVBComponent>(Target, o => new VBComponent((VBAIA.VBComponent) o));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IVBComponent>) this).GetEnumerator();
        }

        public override bool Equals(ISafeComWrapper<VBAIA.VBComponents> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IVBComponents other)
        {
            return Equals(other as SafeComWrapper<VBAIA.VBComponents>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Target);
        }

        public void ImportSourceFile(string path)
        {
            if (IsWrappingNullReference) { return; }

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
                try
                {
                    var temp = this[name];
                }
                catch
                {
                    throw new IndexOutOfRangeException($"Could not find document component named '{name}'.  Try adding a document component with the same name and try again.");
                }

                var component = this[name];
                component.CodeModule.Clear();
                component.CodeModule.AddFromString(codeString);
            }
            else if (ext == ComponentTypeExtensions.FormExtension)
            {
                try
                {
                    var temp = this[name];
                }
                catch
                {
                    Import(path);
                }

                var component = this[name];

                var nonAttributeLines = codeLines.TakeWhile(line => !line.StartsWith("Attribute")).Count();
                var attributeLines = codeLines.Skip(nonAttributeLines).TakeWhile(line => line.StartsWith("Attribute")).Count();
                var declarationsStartLine = nonAttributeLines + attributeLines + 1;
                var correctCodeString = string.Join(Environment.NewLine, codeLines.Skip(declarationsStartLine - 1).ToArray());

                component.CodeModule.Clear();
                component.CodeModule.AddFromString(correctCodeString);
            }
            else if (ext != ComponentTypeExtensions.FormBinaryExtension)
            {
                Import(path);
            }
        }

        public void RemoveSafely(IVBComponent component)
        {
            if (component.IsWrappingNullReference) { return; }

            switch (component.Type)
            {
                case ComponentType.ClassModule:
                case ComponentType.StandardModule:
                case ComponentType.UserForm:
                    Remove(component);
                    break;
                case ComponentType.ActiveXDesigner:
                case ComponentType.Document:
                    component.CodeModule.Clear();
                    break;
                default:
                    break;
            }
        }

        #region Events

        private static void AttachEvents(VBAIA.VBComponents components)
        {
            lock (Locker)
            {
                if (_components == null && components != null)
                {
                    _components = components;
                    _componentAdded = OnComponentAdded;
                    _componentRemoved = OnComponentRemoved;
                    _componentRenamed = OnComponentRenamed;
                    _componentSelected = OnComponentSelected;
                    _componentActivated = OnComponentActivated;
                    _componentReloaded = OnComponentReloaded;
                    ComEventsHelper.Combine(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemAdded, _componentAdded);
                    ComEventsHelper.Combine(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRemoved, _componentRemoved);
                    ComEventsHelper.Combine(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRenamed, _componentRenamed);
                    ComEventsHelper.Combine(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemSelected, _componentSelected);
                    ComEventsHelper.Combine(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemActivated, _componentActivated);
                    ComEventsHelper.Combine(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemReloaded, _componentReloaded);
                }               
            }
        }

        internal static void DetatchEvents()
        {
            lock (Locker)
            {
                if (_components != null)
                {
                    ComEventsHelper.Remove(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemAdded, _componentAdded);
                    ComEventsHelper.Remove(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRemoved, _componentRemoved);
                    ComEventsHelper.Remove(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemRenamed, _componentRenamed);
                    ComEventsHelper.Remove(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemSelected, _componentSelected);
                    ComEventsHelper.Remove(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemActivated, _componentActivated);
                    ComEventsHelper.Remove(_components, VBComponentsEventsGuid, (int)ComponentEventDispId.ItemReloaded, _componentReloaded);
                    _components = null;
                }
            }
        }

        private delegate void ItemAddedDelegate(VBAIA.VBComponent vbComponent);
        private static ItemAddedDelegate _componentAdded;
        public static event EventHandler<ComponentEventArgs> ComponentAdded;
        private static void OnComponentAdded(VBAIA.VBComponent vbComponent)
        {
            OnDispatch(ComponentAdded, vbComponent);
        }

        private delegate void ItemRemovedDelegate(VBAIA.VBComponent vbComponent);
        private static ItemRemovedDelegate _componentRemoved;
        public static event EventHandler<ComponentEventArgs> ComponentRemoved;
        private static void OnComponentRemoved(VBAIA.VBComponent vbComponent)
        {
            OnDispatch(ComponentRemoved, vbComponent);
        }

        private delegate void ItemRenamedDelegate(VBAIA.VBComponent vbComponent, string oldName);
        private static ItemRenamedDelegate _componentRenamed;
        public static event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        private static void OnComponentRenamed(VBAIA.VBComponent vbComponent, string oldName)
        {
            var handler = ComponentRenamed;
            if (handler != null)
            {
                var component = new VBComponent(vbComponent);
                var project = component.Collection.Parent;
                if (project.Protection != ProjectProtection.Locked)
                {
                    handler.Invoke(component, new ComponentRenamedEventArgs(project.ProjectId, project, new VBComponent(vbComponent), oldName));
                }
            }
        }

        private delegate void ItemSelectedDelegate(VBAIA.VBComponent vbComponent);
        private static ItemSelectedDelegate _componentSelected;
        public static event EventHandler<ComponentEventArgs> ComponentSelected;
        private static void OnComponentSelected(VBAIA.VBComponent vbComponent)
        {
            OnDispatch(ComponentSelected, vbComponent);
        }

        private delegate void ItemActivatedDelegate(VBAIA.VBComponent vbComponent);
        private static ItemActivatedDelegate _componentActivated;
        public static event EventHandler<ComponentEventArgs> ComponentActivated;
        private static void OnComponentActivated(VBAIA.VBComponent vbComponent)
        {
            OnDispatch(ComponentActivated, vbComponent);
        }

        private delegate void ItemReloadedDelegate(VBAIA.VBComponent vbComponent);
        private static ItemReloadedDelegate _componentReloaded;
        public static event EventHandler<ComponentEventArgs> ComponentReloaded;
        private static void OnComponentReloaded(VBAIA.VBComponent vbComponent)
        {
            OnDispatch(ComponentReloaded, vbComponent);
        }

        private static void OnDispatch(EventHandler<ComponentEventArgs> dispatched, VBAIA.VBComponent vbComponent)
        {
            var handler = dispatched;
            if (handler != null)
            {
                var component = new VBComponent(vbComponent);
                var project = component.Collection.Parent;
                if (project.Protection != ProjectProtection.Locked)
                {
                    handler.Invoke(component, new ComponentEventArgs(project.ProjectId, project, component));
                }
            }
        }

        #endregion
    }
}