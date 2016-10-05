using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class References : SafeComWrapper<Microsoft.Vbe.Interop.References>, IEnumerable<Reference>, IEquatable<References>
    {
        public References(Microsoft.Vbe.Interop.References comObject) 
            : base(comObject)
        {
            comObject.ItemAdded += comObject_ItemAdded;
            comObject.ItemRemoved += comObject_ItemRemoved;
        }

        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public VBProject Parent
        {
            get { return new VBProject(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Parent)); }
        }

        public VBE VBE
        {
            get { return new VBE(InvokeResult(() => IsWrappingNullReference ? null : ComObject.VBE)); }
        }

        private void comObject_ItemRemoved(Microsoft.Vbe.Interop.Reference reference)
        {
            var handler = ItemRemoved;
            if (handler == null) { return; }
            handler.Invoke(this, new ReferenceEventArgs(new Reference(reference)));
        }

        private void comObject_ItemAdded(Microsoft.Vbe.Interop.Reference reference)
        {
            var handler = ItemAdded;
            if (handler == null) { return; }
            handler.Invoke(this, new ReferenceEventArgs(new Reference(reference)));
        }

        public Reference Item(object index)
        {
            return new Reference(InvokeResult(() => ComObject.Item(index)));
        }

        public Reference AddFromGuid(string guid, int major, int minor)
        {
            return new Reference(InvokeResult(() => ComObject.AddFromGuid(guid, major, minor)));
        }

        public Reference AddFromFile(string path)
        {
            return new Reference(InvokeResult(() => ComObject.AddFromFile(path)));
        }

        public void Remove(Reference reference)
        {
            Invoke(() => ComObject.Remove(reference.ComObject));
        }

        IEnumerator<Reference> IEnumerable<Reference>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Reference>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<Reference>)this).GetEnumerator();
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

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.References> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject.Parent, Parent.ComObject));
        }

        public bool Equals(References other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.References>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}