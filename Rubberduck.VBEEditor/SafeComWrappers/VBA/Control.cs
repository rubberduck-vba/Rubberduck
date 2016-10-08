using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Control : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Control>, IControl
    {
        public Control(Microsoft.Vbe.Interop.Forms.Control comObject) 
            : base(comObject)
        {
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Name; }
            set { ComObject.Name = value; }
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

        public bool Equals(IControl other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Forms.Control>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}