using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Controls : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>, IControls
    {
        public Controls(Microsoft.Vbe.Interop.Forms.Controls comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public IControl this[object index]
        {
            get { return new Control((Microsoft.Vbe.Interop.Forms.Control) ComObject.Item(index)); }
        }

        IEnumerator<IControl> IEnumerable<IControl>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IControl>(ComObject);
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
                Marshal.ReleaseComObject(ComObject);
            } 
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(IControls other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}