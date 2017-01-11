using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AddIn : SafeComWrapper<VB.AddIn>, IAddIn
    {
        public AddIn(Microsoft.Vbe.Interop.AddIn target) 
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
            set { if (!IsWrappingNullReference) Target.Description = value; }
        }

        public bool Connect
        {
            get { return !IsWrappingNullReference && Target.Connect; }
            set { if (!IsWrappingNullReference) Target.Connect = value; }
        }

        public object Object // definitely leaks a COM object
        {
            get { return IsWrappingNullReference ? null : Target.Object; }
            set { if (!IsWrappingNullReference) Target.Object = value; }
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