using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Windows : SafeComWrapper<Microsoft.Vbe.Interop.Windows>, IEnumerable, IEquatable<Windows>
    {
        public Windows(Microsoft.Vbe.Interop.Windows windows)
            : base(windows)
        {
        }

        public int Count
        {
            get { return InvokeResult(() => ComObject.Count); }
        }

        public VBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : InvokeResult(() => ComObject.VBE)); }
        }

        public Application Parent
        {
            get { return new Application(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Parent)); }
        }

        public Window Item(object index)
        {
            return new Window(InvokeResult(() => ComObject.Item(index)));
        }

        public Window CreateToolWindow(AddIn addInInst, string progId, string caption, string guidPosition, ref object docObj)
        {
            try
            {
                return new Window(ComObject.CreateToolWindow(addInInst.ComObject, progId, caption, guidPosition, ref docObj));
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
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

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Windows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(Windows other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Windows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}