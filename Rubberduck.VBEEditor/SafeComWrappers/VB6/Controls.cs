using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    //public class Controls : SafeComWrapper<VB.Controls>, IControls
    //{
    //    public Controls(VB.Controls target) 
    //        : base(target)
    //    {
    //    }

    //    public int Count
    //    {
    //        get { return IsWrappingNullReference ? 0 : Target.Count; }
    //    }

    //    public IControl this[object index]
    //    {
    //        get { return new Control((VB.Control) Target.Item(index)); }
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

    //    public override bool Equals(ISafeComWrapper<VB.Controls> other)
    //    {
    //        return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
    //    }

    //    public bool Equals(IControls other)
    //    {
    //        return Equals(other as SafeComWrapper<VB.Controls>);
    //    }

    //    public override int GetHashCode()
    //    {
    //        return IsWrappingNullReference ? 0 : Target.GetHashCode();
    //    }
    //}
}