using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class LinkedWindows : SafeComWrapper<VB.LinkedWindows>, ILinkedWindows
    {
        public LinkedWindows(VB.LinkedWindows linkedWindows)
            : base(linkedWindows)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IWindow Parent
        {
            get { return new Window(IsWrappingNullReference ? null : Target.Parent); }
        }

        public IWindow this[object index]
        {
            get { return new Window(IsWrappingNullReference ? null : Target.Item(index)); }
        }

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
                : new ComWrapperEnumerator<IWindow>(Target, o => new Window((VB.Window) o));
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
        
        public override bool Equals(ISafeComWrapper<VB.LinkedWindows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ILinkedWindows other)
        {
            return Equals(other as SafeComWrapper<VB.LinkedWindows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}