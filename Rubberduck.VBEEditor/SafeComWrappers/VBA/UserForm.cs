using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class UserForm : SafeComWrapper<Microsoft.Vbe.Interop.Forms.UserForm>, IUserForm
    {
        public UserForm(Microsoft.Vbe.Interop.Forms.UserForm target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public IControls Controls => new Controls(Target.Controls);

        public IControls Selected => new Controls(Target.Selected);

        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.Forms.UserForm> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IUserForm other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Forms.UserForm>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}
