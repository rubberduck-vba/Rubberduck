using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class References : SafeEventedComWrapper<VB.References, VB._dispReferencesEvents>, IReferences, VB._dispReferencesEvents
    {
        public References(VB.References target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;
        
        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBProject Parent => new VBProject(IsWrappingNullReference ? null : Target.Parent);

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        void VB._dispReferencesEvents.ItemRemoved(VB.Reference Reference)
        {
            var referenceWrapper = new Reference(Reference);
            var handler = ItemRemoved;
            if (handler == null)
            {
                referenceWrapper.Dispose();
                return;
            }
            handler.Invoke(this, new ReferenceEventArgs(referenceWrapper));
        }

        void VB._dispReferencesEvents.ItemAdded(VB.Reference Reference)
        {
            var referenceWrapper = new Reference(Reference);
            var handler = ItemAdded;
            if (handler == null)
            {
                referenceWrapper.Dispose();
                return;
            }
            handler.Invoke(this, new ReferenceEventArgs(referenceWrapper));
        }

        public IReference this[object index] => new Reference(IsWrappingNullReference ? null : Target.Item(index));

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
            return new ComWrapperEnumerator<IReference>(Target, comObject => new Reference((VB.Reference) comObject));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IReference>) this).GetEnumerator();
        }

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