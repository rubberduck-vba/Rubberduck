using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class References : SafeComWrapper<VB.References>, IReferences
    {
        public References(VB.References target) 
            : base(target)
        {
        }

        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;

        public void HandleEvents(EventHandler<ReferenceEventArgs> handleItemAdded, EventHandler<ReferenceEventArgs> handleItemRemoved)
        {
            var addedHandler = ItemAdded;
            var removedHandler = ItemRemoved;
            if (addedHandler != null || removedHandler != null)
            {
                return;
            }

            Target.ItemAdded += Target_ItemAdded;
            Target.ItemRemoved += Target_ItemRemoved;
            ItemAdded += handleItemAdded;
            ItemRemoved += handleItemRemoved;
        }

        public void UnregisterEvents(EventHandler<ReferenceEventArgs> handleItemAdded, EventHandler<ReferenceEventArgs> handleItemRemoved)
        {
            var addedHandler = ItemAdded;
            var removedHandler = ItemRemoved;
            if (addedHandler == null || removedHandler == null)
            {
                return;
            }

            Target.ItemAdded -= Target_ItemAdded;
            Target.ItemRemoved -= Target_ItemRemoved;
            ItemAdded -= handleItemAdded;
            ItemRemoved -= handleItemRemoved;
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
            get { return new Reference(IsWrappingNullReference ? null : Target.Item(index)); }
        }

        public IReference AddFromGuid(string guid, int major, int minor)
        {
            return new Reference(IsWrappingNullReference ? null : Target.AddFromGuid(guid, major, minor));
        }

        public IReference AddFromFile(string path)
        {
            return new Reference(IsWrappingNullReference ? null : Target.AddFromFile(path));
        }

        public void Remove(IReference reference)
        {
            if (IsWrappingNullReference) return;
            Target.Remove(((ISafeComWrapper<VB.Reference>)reference).Target);
        }

        IEnumerator<IReference> IEnumerable<IReference>.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<IReference>(null, o => new Reference(null))
                : new ComWrapperEnumerator<IReference>(Target, o => new Reference((VB.Reference) o));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IReference>) this).GetEnumerator();
        }

        //private bool _isReleased;
        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference && !_isReleased)
        //    {
        //        for (var i = 1; i <= Count; i++)
        //        {
        //            this[i].Release();
        //        }

        //        base.Release(final);
        //        _isReleased = true;
        //    }
        //}

        public override bool Equals(ISafeComWrapper<VB.References> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target.Parent, Parent.Target));
        }

        public bool Equals(IReferences other)
        {
            return Equals(other as SafeComWrapper<VB.References>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}