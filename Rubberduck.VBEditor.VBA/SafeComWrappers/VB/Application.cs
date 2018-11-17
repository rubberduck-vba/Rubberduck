using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Application : SafeComWrapper<VB.Application>, IApplication
    {
        public Application(VB.Application target, bool rewrapping = false)
            :base(target, rewrapping)
        {
        }

        public string Version => IsWrappingNullReference ? null : Target.Version;

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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}