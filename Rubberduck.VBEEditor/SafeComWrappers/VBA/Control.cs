using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Control : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Control>, IEquatable<Control>
    {
        public Control(Microsoft.Vbe.Interop.Forms.Control comObject) 
            : base(comObject)
        {
        }

        public string Name
        {
            get { return InvokeResult(() => IsWrappingNullReference ? string.Empty : ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Forms.Control> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(Control other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Forms.Control>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}