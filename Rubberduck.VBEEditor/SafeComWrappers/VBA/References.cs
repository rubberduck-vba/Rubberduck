using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class References : SafeComWrapper<Microsoft.Vbe.Interop.References>, IReferences
    {
        public References(Microsoft.Vbe.Interop.References target) 
            : base(target)
        {
            if (!IsWrappingNullReference)
            {
                target.ItemAdded += Target_ItemAdded;
                target.ItemRemoved += Target_ItemRemoved;
            }
        }

        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;

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

        private void Target_ItemRemoved(Microsoft.Vbe.Interop.Reference reference)
        {
            var handler = ItemRemoved;
            if (handler == null) { return; }
            handler.Invoke(this, new ReferenceEventArgs(new Reference(reference)));
        }

        private void Target_ItemAdded(Microsoft.Vbe.Interop.Reference reference)
        {
            var handler = ItemAdded;
            if (handler == null) { return; }
            handler.Invoke(this, new ReferenceEventArgs(new Reference(reference)));
        }

        public IReference this[object index]
        {
            get { return new Reference(Target.Item(index)); }
        }

        public IReference AddFromGuid(string guid, int major, int minor)
        {
            return new Reference(Target.AddFromGuid(guid, major, minor));
        }

        public IReference AddFromFile(string path)
        {
            return new Reference(Target.AddFromFile(path));
        }

        public void Remove(IReference reference)
        {
            Target.Remove(((ISafeComWrapper<Microsoft.Vbe.Interop.Reference>)reference).Target);
        }

        IEnumerator<IReference> IEnumerable<IReference>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Reference>(Target);
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
                Target.ItemAdded -= Target_ItemAdded;
                Target.ItemRemoved -= Target_ItemRemoved;
                Marshal.ReleaseComObject(Target);
            }
        }

        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.References> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target.Parent, Parent.Target));
        }

        public bool Equals(IReferences other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.References>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}