using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public sealed class CommandBarButtonEvents : SafeComWrapper<VB.CommandBarEvents>, ICommandBarButtonEvents, IEventSource<VB.CommandBarEvents>
    {
        public CommandBarButtonEvents(VB.CommandBarEvents target, bool rewrapping = false)
            : base(target, rewrapping)
        {            
        }

        // Explicit implementation as usage should only be from within a SafeComWrapper
        VB.CommandBarEvents IEventSource<VB.CommandBarEvents>.EventSource => Target;

        public override bool Equals(ISafeComWrapper<VB.CommandBarEvents> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICommandBarButtonEvents other)
        {
            return Equals(other as SafeComWrapper<VB.CommandBarEvents>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}
