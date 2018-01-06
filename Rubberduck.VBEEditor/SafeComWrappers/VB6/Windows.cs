using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class Windows : SafeComWrapper<VB.Windows>, IWindows
    {
        public Windows(VB.Windows target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public int Count => Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IApplication Parent => throw new NotImplementedException();

        public IWindow this[object index] => new Window(Target.Item(index));

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            object control = null;
            var window = new Window(Target.CreateToolWindow((VB.AddIn)addInInst.Target, progId, caption, guidPosition, ref control));
            return new ToolWindowInfo(window, control);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Target.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IWindow>(Target, comObject => new Window((VB.Window)comObject));
        }

        public override bool Equals(ISafeComWrapper<VB.Windows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IWindows other)
        {
            return Equals(other as SafeComWrapper<VB.Windows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}