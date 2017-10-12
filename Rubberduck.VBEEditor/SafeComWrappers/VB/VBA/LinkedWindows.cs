using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VBAIA = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    public class LinkedWindows : SafeComWrapper<VBAIA.LinkedWindows>, ILinkedWindows
    {
        public LinkedWindows(VBAIA.LinkedWindows linkedWindows)
            : base(linkedWindows)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IWindow Parent => new Window(IsWrappingNullReference ? null : Target.Parent);

        public IWindow this[object index] => new Window(IsWrappingNullReference ? null : Target.Item(index));

        public void Remove(IWindow window)
        {
            if (IsWrappingNullReference) return;
            Target.Remove(((Window)window).Target);
        }

        public void Add(IWindow window)
        {
            if (IsWrappingNullReference) return;
            Target.Add(((Window)window).Target);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference ? new List<IEnumerable>().GetEnumerator() : Target.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<IWindow>(null, o => new Window(null))
                : new ComWrapperEnumerator<IWindow>(Target, o => new Window((VBAIA.Window) o));
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        for (var i = 1; i <= Count; i++)
        //        {
        //            this[i].Release();
        //        }
        //        base.Release(final);
        //    }
        //}
        
        public override bool Equals(ISafeComWrapper<VBAIA.LinkedWindows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ILinkedWindows other)
        {
            return Equals(other as SafeComWrapper<VBAIA.LinkedWindows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}