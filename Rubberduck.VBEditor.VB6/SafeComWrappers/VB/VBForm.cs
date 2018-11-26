using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class UserForm : SafeComWrapper<VB.VBForm>, IUserForm
    {
        public UserForm(VB.VBForm target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public IControls Controls => new VBControls(Target.VBControls);

        public IControls Selected => new SelectedVBControls(Target.SelectedVBControls);

        public override bool Equals(ISafeComWrapper<VB.VBForm> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IUserForm other)
        {
            return Equals(other as SafeComWrapper<VB.VBForm>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}
