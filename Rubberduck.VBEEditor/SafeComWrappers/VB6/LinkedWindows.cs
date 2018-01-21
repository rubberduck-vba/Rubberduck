using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class LinkedWindows : SafeComWrapper<VB.LinkedWindows>, ILinkedWindows
    {
        public LinkedWindows(VB.LinkedWindows target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IWindow Parent => new Window(IsWrappingNullReference ? null : Target.Parent);

        public IWindow this[object index] => new Window(Target.Item(index));

        public void Remove(IWindow window)
        {
            Target.Remove(((Window)window).Target);
        }

        public void Add(IWindow window)
        {
            Target.Add(((Window)window).Target);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Target.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IWindow>(Target, o => new Window((VB.Window)o));
        }

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