using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Control : SafeComWrapper<VB.Forms.Control>, IControl
    {
        public Control(Microsoft.Vbe.Interop.Forms.Control target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public string Name
        {
            get => IsWrappingNullReference ? string.Empty : Target.Name;
            set { if (!IsWrappingNullReference) Target.Name = value; }
        }

        public override bool Equals(ISafeComWrapper<VB.Forms.Control> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IControl other)
        {
            return Equals(other as SafeComWrapper<VB.Forms.Control>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}