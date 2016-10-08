using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBComponents : SafeComWrapper<Microsoft.Vbe.Interop.VBComponents>, IVBComponents
    {
        public VBComponents(Microsoft.Vbe.Interop.VBComponents target) 
            : base(target)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public IVBProject Parent
        {
            get { return new VBProject((Microsoft.Vbe.Interop.VBProject)(IsWrappingNullReference ? null : Target.Parent.Target)); }
        }

        public IVBE VBE
        {
            get { return new VBE((Microsoft.Vbe.Interop.VBE) (IsWrappingNullReference ? null : Target.VBE.Target)); }
        }

        public IVBComponent this[object index]
        {
            get { return new VBComponent((Microsoft.Vbe.Interop.VBComponent)(IsWrappingNullReference ? null : Target[index].Target)); }
        }

        public void Remove(IVBComponent item)
        {
            Target.Remove(item);
        }

        public IVBComponent Add(ComponentType type)
        {
            return new VBComponent((Microsoft.Vbe.Interop.VBComponent) Target.Add(type).Target);
        }

        public IVBComponent Import(string path)
        {
            return new VBComponent((Microsoft.Vbe.Interop.VBComponent) Target.Import(path).Target);
        }

        public IVBComponent AddCustom(string progId)
        {
            return new VBComponent((Microsoft.Vbe.Interop.VBComponent) Target.AddCustom(progId).Target);
        }

        public IVBComponent AddMTDesigner(int index = 0)
        {
            return new VBComponent((Microsoft.Vbe.Interop.VBComponent) Target.AddMTDesigner(index).Target);
        }

        IEnumerator<IVBComponent> IEnumerable<IVBComponent>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IVBComponent>(Target);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IVBComponent>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(Target);
            }
        }

        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.VBComponents> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IVBComponents other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBComponents>);
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

        public IVBComponents Target { get; private set; }
    }
}