using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class LinkedWindows : SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows>, ILinkedWindows
    {
        public LinkedWindows(Microsoft.Vbe.Interop.LinkedWindows linkedWindows)
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
            get { return new Window(Target.Item(index)); }
        }

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
            return new ComWrapperEnumerator<IWindow>(Target);
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(Target);
            }
        }
        
        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ILinkedWindows other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}