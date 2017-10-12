using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VB6
{
    public class Properties : SafeComWrapper<VB6IA.Properties>, IProperties
    {
        public Properties(VB6IA.Properties target) 
            : base(target)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IApplication Application
        {
            get { return new Application((VB6IA.Application) (IsWrappingNullReference ? null : Target.Application)); }
        }

        public object Parent
        {
            get { return IsWrappingNullReference ? null : Target.Parent; }
        }

        public IProperty this[object index]
        {
            get { return new Property(Target.Item(index)); }
        }

        IEnumerator<IProperty> IEnumerable<IProperty>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IProperty>(Target, o => new Property((VB6IA.Property)o));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IProperty>)this).GetEnumerator();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        for (var i = 1; i <= Count; i++)
        //        {
        //            this[i].Release();
        //        }
        //        base.Release(final);
        //    }
        //}

        public override bool Equals(ISafeComWrapper<VB6IA.Properties> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IProperties other)
        {
            return Equals(other as SafeComWrapper<VB6IA.Properties>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}