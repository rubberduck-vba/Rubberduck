using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class LinkedWindows : SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows>, IEnumerable, IEquatable<LinkedWindows>
    {
        public LinkedWindows(Microsoft.Vbe.Interop.LinkedWindows linkedWindows)
            : base(linkedWindows)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public VBE VBE
        {
            get { return new VBE(InvokeResult(() => IsWrappingNullReference ? null : ComObject.VBE)); }
        }

        public Window Parent
        {
            get { return new Window(InvokeResult(() => IsWrappingNullReference ? null : ComObject.Parent)); }
        }

        public Window Item(object index)
        {
            return new Window(InvokeResult(() => ComObject.Item(index)));
        }

        public void Remove(Window window)
        {
            Invoke(() => ComObject.Remove(window.ComObject));
        }

        public void Add(Window window)
        {
            Invoke(() => ComObject.Add(window.ComObject));
        }

        public IEnumerator GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    Item(i).Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(LinkedWindows other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}