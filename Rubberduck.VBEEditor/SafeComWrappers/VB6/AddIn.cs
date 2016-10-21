using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class AddIn : SafeComWrapper<VB.AddIn>, IAddIn
    {
        public AddIn(VB.AddIn target) 
            : base(target)
        {
        }

        public string ProgId
        {
            get { return IsWrappingNullReference ? string.Empty : Target.ProgId; }
        }

        public string Guid
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Guid; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IAddIns Collection
        {
            get { return new AddIns(IsWrappingNullReference ? null : Target.Collection); }
        }

        public string Description
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Description; } 
            set { Target.Description = value; }
        }

        public bool Connect
        {
            get { return !IsWrappingNullReference && Target.Connect; }
            set { Target.Connect = value; }
        }

        public object Object // definitely leaks a COM object
        {
            get { return IsWrappingNullReference ? null : Target.Object; }
            set { Target.Object = value; }
        }

        public override bool Equals(ISafeComWrapper<VB.AddIn> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.ProgId == ProgId && other.Target.Guid == Guid);
        }

        public bool Equals(IAddIn other)
        {
            return Equals(other as SafeComWrapper<VB.AddIn>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(ProgId, Guid);
        }
    }
}