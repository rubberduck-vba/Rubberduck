using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Controls : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>, IEnumerable<Control>, IEquatable<Controls>
    {
        public Controls(Microsoft.Vbe.Interop.Forms.Controls comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public Control Item(object index)
        {
            return new Control(InvokeResult(() => (Microsoft.Vbe.Interop.Forms.Control)ComObject.Item(index)));
        }

        IEnumerator<Control> IEnumerable<Control>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Control>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<Control>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    Item(i).Release();
                }
                Marshal.ReleaseComObject(ComObject);
            } 
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls> other)
        {
            return IsEqualIfNull(other) || ReferenceEquals(other.ComObject, ComObject);
        }

        public bool Equals(Controls other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}