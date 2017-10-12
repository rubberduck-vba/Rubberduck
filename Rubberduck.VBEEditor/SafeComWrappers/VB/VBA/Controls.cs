using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Office.Abstract;
using VBAIA = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    public class Controls : SafeComWrapper<VBAIA.Forms.Controls>, IControls
    {
        public Controls(VBAIA.Forms.Controls target) 
            : base(target)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IControl this[object index] => IsWrappingNullReference ? new Control(null) : new Control((VBAIA.Forms.Control) Target.Item(index));

        IEnumerator<IControl> IEnumerable<IControl>.GetEnumerator()
        {
            // soft-casting because ImageClass doesn't implement IControl
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<IControl>(null, o => new Control(null))
                : new ComWrapperEnumerator<IControl>(Target, o => new Control(o as VBAIA.Forms.Control));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IControl>) this).GetEnumerator();
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        //for (var i = 1; i <= Count; i++)
        //        //{
        //        //    this[i].Release();
        //        //}
        //        base.Release(final);
        //    } 
        //}

        public override bool Equals(ISafeComWrapper<VBAIA.Forms.Controls> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IControls other)
        {
            return Equals(other as SafeComWrapper<VBAIA.Forms.Controls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}