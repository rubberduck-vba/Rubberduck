using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Application : SafeComWrapper<Microsoft.Vbe.Interop.Application>, IApplication
    {
        public Application(Microsoft.Vbe.Interop.Application application)
            :base(application)
        {
        }

        public string Version { get { return IsWrappingNullReference ? null : Target.Version; } }
        
        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                Marshal.ReleaseComObject(Target);
            }
        }

        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.Application> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.Version == Version);
        }

        public bool Equals(IApplication other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Application>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Target);
        }
    }
}