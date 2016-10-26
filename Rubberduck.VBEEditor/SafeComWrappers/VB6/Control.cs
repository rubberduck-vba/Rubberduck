using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    //public class Control : SafeComWrapper<VB.Control>, IControl
    //{
    //    public Control(Microsoft.Vbe.Interop.Forms.Control target) 
    //        : base(target)
    //    {
    //    }

    //    public string Name
    //    {
    //        get { return IsWrappingNullReference ? string.Empty : Target.Name; }
    //        set { Target.Name = value; }
    //    }

    //    public override void Release()
    //    {
    //        if (!IsWrappingNullReference)
    //        {
    //            Marshal.ReleaseComObject(Target);
    //        }
    //    }

    //    public override bool Equals(ISafeComWrapper<VB.Control> other)
    //    {
    //        return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
    //    }

    //    public bool Equals(IControl other)
    //    {
    //        return Equals(other as SafeComWrapper<VB.Control>);
    //    }

    //    public override int GetHashCode()
    //    {
    //        return IsWrappingNullReference ? 0 : Target.GetHashCode();
    //    }
    //}
}