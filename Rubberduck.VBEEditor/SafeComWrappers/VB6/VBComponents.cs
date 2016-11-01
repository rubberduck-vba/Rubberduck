using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBComponents : SafeComWrapper<VB.VBComponents>, IVBComponents
    {
        private readonly IVBComponentsEventsSink _sink;

        public VBComponents(VB.VBComponents target) 
            : base(target)
        {
            _sink = new VBComponentsEventsSink();
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public IVBProject Parent
        {
            get { return new VBProject(IsWrappingNullReference ? null : Target.Parent); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IVBComponent this[object index]
        {
            get { return new VBComponent(IsWrappingNullReference ? null : Target.Item(index)); }
        }

        public void Remove(IVBComponent item)
        {
            Target.Remove((VB.VBComponent)item.Target);
        }

        public IVBComponent Add(ComponentType type)
        {
            return new VBComponent(Target.Add((VB.vbext_ComponentType)type));
        }

        public IVBComponent Import(string path)
        {
            throw new NotImplementedException();
            //return new VBComponent((VB.VBComponent)Target.Import(path).Target);
        }

        public IVBComponent AddCustom(string progId)
        {
            throw new NotImplementedException();
            //return new VBComponent((VB.VBComponent)Target.AddCustom(progId).Target);
        }

        public IVBComponent AddMTDesigner(int index = 0)
        {
            throw new NotImplementedException();
            //return new VBComponent((VB.VBComponent)Target.AddMTDesigner(index).Target);
        }

        IEnumerator<IVBComponent> IEnumerable<IVBComponent>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IVBComponent>(Target, o => new VBComponent((VB.VBComponent)o));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IVBComponent>)this).GetEnumerator();
        }

        public override void Release(bool final = false)
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                base.Release(final);
            }
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
                component.CodeModule.Clear();
                component.CodeModule.AddFromString(codeString);
            }
            else if (ext == ComponentTypeExtensions.FormExtension)
            {
                IVBComponent component;
                try
                {
                    component = this[name];
                }
                catch
                {
                    component = Add(ComponentType.UserForm);
                    component.Properties["Caption"].Value = name;
                    component.Name = name;
                }

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
                    component.CodeModule.Clear();
                    break;
                default:
                    break;
            }
        }

        public IVBComponentsEventsSink Events { get { return _sink; } }

        public IConnectionPoint ConnectionPoint
        {
            get
            {
                try
                {
                    // ReSharper disable once SuspiciousTypeConversion.Global
                    var connectionPointContainer = (IConnectionPointContainer)Target;
                    var interfaceId = typeof(VB._dispVBComponentsEvents).GUID;
                    IConnectionPoint result;
                    connectionPointContainer.FindConnectionPoint(ref interfaceId, out result);
                    return result;
                }
                catch (Exception)
                {
                    return null;
                }

            }
        }
    }
}