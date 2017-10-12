using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VBAIA = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    public class Application : SafeComWrapper<VBAIA.Application>, IApplication
    {
        public Application(VBAIA.Application application)
            :base(application)
        {
        }

        public string Version { get { return IsWrappingNullReference ? null : Target.Version; } }

        public override bool Equals(ISafeComWrapper<VBAIA.Application> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.Version == Version);
        }

        public bool Equals(IApplication other)
        {
            return Equals(other as SafeComWrapper<VBAIA.Application>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Target);
        }
    }
}