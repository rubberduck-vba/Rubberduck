using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBControl : SafeComWrapper<VB.VBControl>, IControl
    {
        public VBControl(VB.VBControl target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public string Name
        {
            get => IsWrappingNullReference ? string.Empty : Target.Properties.Item("Name").Value.ToString();
            set { if (!IsWrappingNullReference) Target.Properties.Item("Name").Value = value; }
        }

        public override bool Equals(ISafeComWrapper<VB.VBControl> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IControl other)
        {
            return Equals(other as SafeComWrapper<VB.VBControl>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}