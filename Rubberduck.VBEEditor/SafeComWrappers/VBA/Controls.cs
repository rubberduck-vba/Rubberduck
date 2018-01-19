using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Controls : SafeComWrapper<VB.Forms.Controls>, IControls
    {
        public Controls(VB.Forms.Controls target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IControl this[object index] => IsWrappingNullReference ? new Control(null) : new Control((VB.Forms.Control) Target.Item(index));

        IEnumerator<IControl> IEnumerable<IControl>.GetEnumerator()
        {
            // soft-casting because ImageClass doesn't implement IControl
            return new ComWrapperEnumerator<IControl>(Target, comObject => new Control(comObject as VB.Forms.Control));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IControl>) this).GetEnumerator();
        }

        public override bool Equals(ISafeComWrapper<VB.Forms.Controls> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IControls other)
        {
            return Equals(other as SafeComWrapper<VB.Forms.Controls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}