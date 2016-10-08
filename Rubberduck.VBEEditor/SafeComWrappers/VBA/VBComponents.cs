using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBComponents : SafeComWrapper<Microsoft.Vbe.Interop.VBComponents>, IEnumerable<VBComponent>, IEquatable<VBComponents>
    {
        public VBComponents(Microsoft.Vbe.Interop.VBComponents comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public VBProject Parent
        {
            get { return new VBProject(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public VBComponent this[object index]
        {
            get { return new VBComponent(IsWrappingNullReference ? null : ComObject.Item(index)); }
        }

        public void Remove(VBComponent item)
        {
            ComObject.Remove(item.ComObject);
        }

        public VBComponent Add(ComponentType type)
        {
            return new VBComponent(ComObject.Add((vbext_ComponentType)type));
        }

        public VBComponent Import(string path)
        {
            return new VBComponent(ComObject.Import(path));
        }

        public VBComponent AddCustom(string progId)
        {
            return new VBComponent(ComObject.AddCustom(progId));
        }

        public VBComponent AddMTDesigner(int index = 0)
        {
            return new VBComponent(ComObject.AddMTDesigner(index));
        }

        IEnumerator<VBComponent> IEnumerable<VBComponent>.GetEnumerator()
        {
            return new ComWrapperEnumerator<VBComponent>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<VBComponent>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.VBComponents> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(VBComponents other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBComponents>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
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
                VBComponent component;
                try
                {
                    component = this[name];
                }
                catch (IndexOutOfRangeException)
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
        public void RemoveSafely(VBComponent component)
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
    }
}