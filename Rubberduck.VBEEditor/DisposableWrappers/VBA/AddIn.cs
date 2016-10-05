using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class AddIn : SafeComWrapper<Microsoft.Vbe.Interop.AddIn>, IEquatable<AddIn>
    {
        public AddIn(Microsoft.Vbe.Interop.AddIn comObject) 
            : base(comObject)
        {
        }

        public string ProgId
        {
            get
            {
                return IsWrappingNullReference ? null : InvokeResult(() => ComObject.ProgId);
            }
        }

        public string Guid
        {
            get { return IsWrappingNullReference ? null : InvokeResult(() => ComObject.Guid); }
        }

        public VBE VBE
        {
            get { return new VBE(InvokeResult(() => IsWrappingNullReference ? null : ComObject.VBE)); }
        }

        public AddIns Collection
        {
            get { return new AddIns(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Collection)); }
        }

        public string Description
        {
            get
            {
                return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Description);
            }
            set
            {
                Invoke(() => ComObject.Description = value);
            }
        }

        public bool Connect
        {
            get
            {
                return !IsWrappingNullReference && InvokeResult(() => ComObject.Connect);
            }
            set
            {
                Invoke(() => ComObject.Connect = value);
            }
        }

        public object Object // definitely leaks a COM object
        {
            get
            {
                return IsWrappingNullReference ? null : InvokeResult(() => ComObject.Object);
            }
            set
            {
                Invoke(() => ComObject.Object = value);
            }
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
            return IsEqualIfNull(other) || (other.ComObject.ProgId == ProgId && other.ComObject.Guid == Guid);
        }

        public bool Equals(AddIn other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.AddIn>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(ProgId, Guid);
        }
    }
}