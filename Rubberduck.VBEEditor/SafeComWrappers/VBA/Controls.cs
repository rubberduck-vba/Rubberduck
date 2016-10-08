using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Controls : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>, IControls
    {
        public Controls(Microsoft.Vbe.Interop.Forms.Controls target) 
            : base(target)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public IControl this[object index]
        {
            get { return new Control((Microsoft.Vbe.Interop.Forms.Control) Target.Item(index)); }
        }

        IEnumerator<IControl> IEnumerable<IControl>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IControl>(Target);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IControl>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(Target);
            } 
        }

        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IControls other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}