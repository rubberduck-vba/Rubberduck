using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public sealed class AddIns : SafeComWrapper<VB.Addins>, IAddIns
    {
        public AddIns(VB.Addins target, bool rewrapping = false) : 
            base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public object Parent // todo: verify if this could be 'public Application Parent' instead
            => IsWrappingNullReference ? null : Target.Parent;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IAddIn this[object index] => new AddIn(IsWrappingNullReference ? null : Target.Item(index));

        public void Update()
        {
            Target.Update();
        }

        public override bool Equals(ISafeComWrapper<VB.Addins> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target.Parent, Parent));
        }

        public bool Equals(IAddIns other)
        {
            return Equals(other as ISafeComWrapper<VB.Addins>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Parent);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator)new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IAddIn>)this).GetEnumerator();
        }

        IEnumerator<IAddIn> IEnumerable<IAddIn>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IAddIn>(Target, comObject => new AddIn((VB.AddIn) comObject));
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}