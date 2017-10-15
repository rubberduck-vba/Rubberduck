using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Application : SafeComWrapper<VB.Application>, IApplication
    {
        public Application(VB.Application application)
            :base(application)
        {
        }

        public string Version { get { return IsWrappingNullReference ? null : Target.Version; } }

        public override bool Equals(ISafeComWrapper<VB.Application> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.Version == Version);
        }

        public bool Equals(IApplication other)
        {
            return Equals(other as SafeComWrapper<VB.Application>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Target);
        }
    }
}