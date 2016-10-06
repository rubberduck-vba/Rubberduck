using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public IWindow Parent
        {
            get { return new Window(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public IWindow this[object index]
        {
            get { return new Window(ComObject.Item(index)); }
        }

        public void Remove(IWindow window)
        {
            ComObject.Remove(((Window)window).ComObject);
        }

        public void Add(IWindow window)
        {
            ComObject.Add(((Window)window).ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ComObject.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IWindow>(ComObject);
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }
        
        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(ILinkedWindows other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}