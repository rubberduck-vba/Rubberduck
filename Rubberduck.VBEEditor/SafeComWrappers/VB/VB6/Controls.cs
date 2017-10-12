using VBIA = Microsoft.VB6.Interop.VBIDE;
namespace Rubberduck.VBEditor.SafeComWrappers.VB.VB6
{
    //public class Controls : SafeComWrapper<VB6IA.Controls>, IControls
    //{
    //    public Controls(VB6IA.Controls target) 
    //        : base(target)
    //    {
    //    }

    //    public int Count
    //    {
    //        get { return IsWrappingNullReference ? 0 : Target.Count; }
    //    }

    //    public IControl this[object index]
    //    {
    //        get { return new Control((VB6IA.Control) Target.Item(index)); }
    //    }

    //    IEnumerator<IControl> IEnumerable<IControl>.GetEnumerator()
    //    {
    //        return new ComWrapperEnumerator<IControl>(Target);
    //    }

    //    IEnumerator IEnumerable.GetEnumerator()
    //    {
    //        return ((IEnumerable<IControl>)this).GetEnumerator();
    //    }

    //    public override void Release()
    //    {
    //        if (!IsWrappingNullReference)
    //        {
    //            for (var i = 1; i <= Count; i++)
    //            {
    //                this[i].Release();
    //            }
    //            Marshal.ReleaseComObject(Target);
    //        } 
    //    }

    //    public override bool Equals(ISafeComWrapper<VB6IA.Controls> other)
    //    {
    //        return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
    //    }

    //    public bool Equals(IControls other)
    //    {
    //        return Equals(other as SafeComWrapper<VB6IA.Controls>);
    //    }

    //    public override int GetHashCode()
    //    {
    //        return IsWrappingNullReference ? 0 : Target.GetHashCode();
    //    }
    //}
}