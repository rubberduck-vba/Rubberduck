using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class Events : SafeComWrapper<VB.Events>, IEvents
    {
        public Events(VB.Events target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public ICommandBarEvents CommandBarEvents => new CommandBarEvents(IsWrappingNullReference ? null : Target, true);

        public override bool Equals(ISafeComWrapper<VB.Events> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IEvents other)
        {
            return Equals(other as SafeComWrapper<VB.Events>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}