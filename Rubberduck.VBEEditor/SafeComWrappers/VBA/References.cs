using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class References : SafeComWrapper<Microsoft.Vbe.Interop.References>, IReferences
    {
        public References(Microsoft.Vbe.Interop.References comObject) 
            : base(comObject)
        {
            if (!IsWrappingNullReference)
            {
                comObject.ItemAdded += comObject_ItemAdded;
                comObject.ItemRemoved += comObject_ItemRemoved;
            }
        }

        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;

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

        public IReference this[object index]
        {
            get { return new Reference(ComObject.Item(index)); }
        }

        public IReference AddFromGuid(string guid, int major, int minor)
        {
            return new Reference(ComObject.AddFromGuid(guid, major, minor));
        }

        public IReference AddFromFile(string path)
        {
            return new Reference(ComObject.AddFromFile(path));
        }

        public void Remove(IReference reference)
        {
            ComObject.Remove((Microsoft.Vbe.Interop.Reference)reference.ComObject);
        }

        IEnumerator<IReference> IEnumerable<IReference>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Reference>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IReference>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                ComObject.ItemAdded -= comObject_ItemAdded;
                ComObject.ItemRemoved -= comObject_ItemRemoved;
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.References> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject.Parent, Parent.ComObject));
        }

        public bool Equals(IReferences other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.References>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}