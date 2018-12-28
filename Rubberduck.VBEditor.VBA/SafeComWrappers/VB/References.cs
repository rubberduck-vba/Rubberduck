using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public sealed class References : SafeEventedComWrapper<VB.References, VB._dispReferences_Events>, IReferences, VB._dispReferences_Events
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

        void VB._dispReferences_Events.ItemRemoved(VB.Reference reference)
        {
            using (var removing = new Reference(reference))
            {
                ItemRemoved?.Invoke(this, new ReferenceEventArgs(new ReferenceInfo(removing), removing.Type));
            }
        }

        void VB._dispReferences_Events.ItemAdded(VB.Reference reference)
        {
            using (var adding = new Reference(reference))
            {
                ItemAdded?.Invoke(this, new ReferenceEventArgs(new ReferenceInfo(adding), adding.Type));
            }
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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}