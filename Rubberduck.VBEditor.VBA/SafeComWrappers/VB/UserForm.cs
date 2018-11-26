using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class UserForm : SafeComWrapper<VB.Forms.UserForm>, IUserForm
    {
        public UserForm(Microsoft.Vbe.Interop.Forms.UserForm target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public IControls Controls => new Controls(Target.Controls);

        public IControls Selected => new Controls(Target.Selected);

        public override bool Equals(ISafeComWrapper<VB.Forms.UserForm> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IUserForm other)
        {
            return Equals(other as SafeComWrapper<VB.Forms.UserForm>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}
