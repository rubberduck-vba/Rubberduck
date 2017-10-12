using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Abstract;
using VBAIA = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    public class Control : SafeComWrapper<VBAIA.Forms.Control>, IControl
    {
        public Control(VBAIA.Forms.Control target) 
            : base(target)
        {
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Name; }
            set { if (!IsWrappingNullReference) Target.Name = value; }
        }

        public override bool Equals(ISafeComWrapper<VBAIA.Forms.Control> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IControl other)
        {
            return Equals(other as SafeComWrapper<VBAIA.Forms.Control>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}