using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VB6
{
    public class Application : SafeComWrapper<VB6IA.Application>, IApplication
    {
        public Application(VB6IA.Application application)
            :base(application)
        {
        }

        public string Version { get { return IsWrappingNullReference ? null : Target.Version; } }

        public override bool Equals(ISafeComWrapper<VB6IA.Application> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.Version == Version);
        }

        public bool Equals(IApplication other)
        {
            return Equals(other as SafeComWrapper<VB6IA.Application>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(Target);
        }
    }
}