using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Reference : SafeComWrapper<Microsoft.Vbe.Interop.Reference>, IEquatable<Reference>
    {
        public Reference(Microsoft.Vbe.Interop.Reference comObject) 
            : base(comObject)
        {
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Name); }
        }

        public string Guid
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Guid); }
        }

        public int Major
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Major); }
        }

        public int Minor
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Minor); }
        }

        public string Description
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Description); }
        }

        public string FullPath
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.FullPath); }
        }

        public bool IsBuiltIn
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.BuiltIn); }
        }

        public bool IsBroken
        {
            get { return IsWrappingNullReference || InvokeResult(() => ComObject.IsBroken); }
        }

        public ReferenceKind Type
        {
            get { return IsWrappingNullReference ? 0 : (ReferenceKind)InvokeResult(() => ComObject.Type); }
        }

        public References Collection
        {
            get { return new References(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Collection)); }
        }

        public VBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : InvokeResult(() => ComObject.VBE)); }
        }

        public override void Release()
        {
            Marshal.ReleaseComObject(ComObject);
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Reference> other)
        {
            return IsEqualIfNull(other) ||
                   (other != null 
                    && (int)other.ComObject.Type == (int)Type
                    && other.ComObject.Name == Name
                    && other.ComObject.Guid == Guid
                    && other.ComObject.FullPath == FullPath
                    && other.ComObject.Major == Major
                    && other.ComObject.Minor == Minor);
        }

        public bool Equals(Reference other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Reference>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(Type, Name, Guid, FullPath, Major, Minor);
        }
    }
}