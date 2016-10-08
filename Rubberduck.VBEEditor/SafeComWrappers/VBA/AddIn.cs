using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AddIn : SafeComWrapper<Microsoft.Vbe.Interop.AddIn>, ISafeComWrapper, IAddIn
    {
        public AddIn(Microsoft.Vbe.Interop.AddIn comObject) 
            : base(comObject)
        {
        }

        public string ProgId
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.ProgId; }
        }

        public string Guid
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Guid; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public IAddIns Collection
        {
            get { return new AddIns(IsWrappingNullReference ? null : ComObject.Collection); }
        }

        public string Description
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Description; } 
            set { ComObject.Description = value; }
        }

        public bool Connect
        {
            get { return !IsWrappingNullReference && ComObject.Connect; }
            set { ComObject.Connect = value; }
        }

        public object Object // definitely leaks a COM object
        {
            get { return IsWrappingNullReference ? null : ComObject.Object; }
            set { ComObject.Object = value; }
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.AddIn> other)
        {
            return IsEqualIfNull(other) || (other != null && other.ComObject.ProgId == ProgId && other.ComObject.Guid == Guid);
        }

        public bool Equals(IAddIn other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.AddIn>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(ProgId, Guid);
        }
    }
}