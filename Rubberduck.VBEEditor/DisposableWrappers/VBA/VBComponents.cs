using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class VBComponents : SafeComWrapper<Microsoft.Vbe.Interop.VBComponents>, IEnumerable<VBComponent>, IEquatable<VBComponents>
    {
        public VBComponents(Microsoft.Vbe.Interop.VBComponents comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public VBProject Parent
        {
            get { return new VBProject(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Parent)); }
        }

        public VBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : InvokeResult(() => ComObject.VBE)); }
        }

        public VBComponent Item(object index)
        {
            return new VBComponent(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Item(index)));
        }

        public void Remove(VBComponent item)
        {
            Invoke(() => ComObject.Remove(item.ComObject));
        }

        public VBComponent Add(ComponentType type)
        {
            return new VBComponent(InvokeResult(() => ComObject.Add((vbext_ComponentType)type)));
        }

        public VBComponent Import(string path)
        {
            return new VBComponent(InvokeResult(() => ComObject.Import(path)));
        }

        public VBComponent AddCustom(string progId)
        {
            return new VBComponent(InvokeResult(() => ComObject.AddCustom(progId)));
        }

        public VBComponent AddMTDesigner(int index = 0)
        {
            return new VBComponent(InvokeResult(() => ComObject.AddMTDesigner(index)));
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
                    Item(i).Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.VBComponents> other)
        {
            return IsEqualIfNull(other) || ReferenceEquals(other.ComObject, ComObject);
        }

        public bool Equals(VBComponents other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBComponents>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}