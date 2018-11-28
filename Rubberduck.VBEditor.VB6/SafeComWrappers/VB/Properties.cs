using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public sealed class Properties : SafeComWrapper<VB.Properties>, IProperties
    {
        public Properties(VB.Properties target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IApplication Application => new Application(IsWrappingNullReference ? null : Target.Application);

        public object Parent => IsWrappingNullReference ? null : Target.Parent;

        public IProperty this[object index] => new Property(IsWrappingNullReference ? null : Target.Item(index));

        IEnumerator<IProperty> IEnumerable<IProperty>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IProperty>(Target, comObject => new Property((VB.Property) comObject));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IProperty>) this).GetEnumerator();
        }

        public override bool Equals(ISafeComWrapper<VB.Properties> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IProperties other)
        {
            return Equals(other as SafeComWrapper<VB.Properties>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}