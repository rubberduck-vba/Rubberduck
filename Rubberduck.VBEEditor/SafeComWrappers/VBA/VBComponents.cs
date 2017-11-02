using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBComponents : SafeComWrapper<VB.VBComponents>, IVBComponents
    {
        private static readonly Guid VBComponentsEventsGuid = new Guid("0002E116-0000-0000-C000-000000000046");
        private static readonly object Locker = new object();
        private static VB.VBComponents _components;

        private enum ComponentEventDispId
        {
            ItemAdded = 1,
            ItemRemoved = 2,
            ItemRenamed = 3,
            ItemSelected = 4,
            ItemActivated = 5,
            ItemReloaded = 6
        }

        public VBComponents(VB.VBComponents target) : base(target)
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
                Target.Remove((VB.VBComponent)item.Target);
            }
        }

        public IVBComponent Add(ComponentType type)
        {
            return new VBComponent(IsWrappingNullReference ? null : Target.Add((VB.vbext_ComponentType)type));
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
                : new ComWrapperEnumerator<IVBComponent>(Target, o => new VBComponent((VB.VBComponent) o));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IVBComponent>) this).GetEnumerator();
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
            if (IsWrappingNullReference) { return; }

            var ext = Path.GetExtension(path);
            var name = Path.GetFileNameWithoutExtension(path);
            if (!File.Exists(path))
            {
                return;
            }

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

                var codeString = File.ReadAllText(path, Encoding.UTF8);
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

                var codeString = File.ReadAllText(path, Encoding.Default);  //The VBE uses the current ANSI codepage from the windows settings to export and import.
                var codeLines = codeString.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

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

        private static void AttachEvents(VB.VBComponents components)
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

        private delegate void ItemAddedDelegate(VB.VBComponent vbComponent);
        private static ItemAddedDelegate _componentAdded;
        public static event EventHandler<ComponentEventArgs> ComponentAdded;
        private static void OnComponentAdded(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentAdded, vbComponent);
        }

        private delegate void ItemRemovedDelegate(VB.VBComponent vbComponent);
        private static ItemRemovedDelegate _componentRemoved;
        public static event EventHandler<ComponentEventArgs> ComponentRemoved;
        private static void OnComponentRemoved(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentRemoved, vbComponent);
        }

        private delegate void ItemRenamedDelegate(VB.VBComponent vbComponent, string oldName);
        private static ItemRenamedDelegate _componentRenamed;
        public static event EventHandler<ComponentRenamedEventArgs> ComponentRenamed;
        private static void OnComponentRenamed(VB.VBComponent vbComponent, string oldName)
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

        private delegate void ItemSelectedDelegate(VB.VBComponent vbComponent);
        private static ItemSelectedDelegate _componentSelected;
        public static event EventHandler<ComponentEventArgs> ComponentSelected;
        private static void OnComponentSelected(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentSelected, vbComponent);
        }

        private delegate void ItemActivatedDelegate(VB.VBComponent vbComponent);
        private static ItemActivatedDelegate _componentActivated;
        public static event EventHandler<ComponentEventArgs> ComponentActivated;
        private static void OnComponentActivated(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentActivated, vbComponent);
        }

        private delegate void ItemReloadedDelegate(VB.VBComponent vbComponent);
        private static ItemReloadedDelegate _componentReloaded;
        public static event EventHandler<ComponentEventArgs> ComponentReloaded;
        private static void OnComponentReloaded(VB.VBComponent vbComponent)
        {
            OnDispatch(ComponentReloaded, vbComponent);
        }

        private static void OnDispatch(EventHandler<ComponentEventArgs> dispatched, VB.VBComponent vbComponent)
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