using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class Properties : SafeComWrapper<VB.Properties>, IProperties
    {
        public Properties(VB.Properties target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IApplication Application => new Application((VB.Application) (IsWrappingNullReference ? null : Target.Application));

        public object Parent => IsWrappingNullReference ? null : Target.Parent;

        public IProperty this[object index] => new Property(Target.Item(index));

        IEnumerator<IProperty> IEnumerable<IProperty>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IProperty>(Target, comObject => new Property((VB.Property)comObject));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IProperty>)this).GetEnumerator();
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
    }
}