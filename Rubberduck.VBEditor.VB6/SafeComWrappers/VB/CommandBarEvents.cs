using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class CommandBarEvents : SafeComWrapper<VB.Events>, ICommandBarEvents
    {
        public CommandBarEvents(VB.Events target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public ICommandBarButtonEvents this[object button] => new CommandBarButtonEvents(IsWrappingNullReference ? null : Target.CommandBarEvents[button]);

        public override bool Equals(ISafeComWrapper<VB.Events> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICommandBarEvents other)
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