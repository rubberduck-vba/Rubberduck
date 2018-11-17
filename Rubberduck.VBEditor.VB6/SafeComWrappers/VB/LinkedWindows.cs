using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public sealed class LinkedWindows : SafeComWrapper<VB.LinkedWindows>, ILinkedWindows
    {
        public LinkedWindows(VB.LinkedWindows target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IWindow Parent => new Window(IsWrappingNullReference ? null : Target.Parent);

        public IWindow this[object index] => new Window(IsWrappingNullReference ? null : Target.Item(index));

        public void Remove(IWindow window)
        {
            if (IsWrappingNullReference)
            {
                return;
            }

            Target.Remove(((Window)window).Target);
        }

        public void Add(IWindow window)
        {
            if (IsWrappingNullReference)
            {
                return;
            }

            Target.Add(((Window)window).Target);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator)new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IWindow>)this).GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IWindow>(Target, comObject => new Window((VB.Window) comObject));
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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}