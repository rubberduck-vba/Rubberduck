using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Application : SafeComWrapper<Microsoft.Vbe.Interop.Application>, IEquatable<Application>
    {
        public Application(Microsoft.Vbe.Interop.Application application)
            :base(application)
        {
        }

        public string Version { get { return IsWrappingNullReference ? null : InvokeResult(() => ComObject.Version); } }
        
        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Application> other)
        {
            return IsEqualIfNull(other) || other.ComObject.Version == Version;
        }

        public bool Equals(Application other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Application>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}