using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Windows : SafeComWrapper<Microsoft.Vbe.Interop.Windows>, IWindows
    {
        public Windows(Microsoft.Vbe.Interop.Windows windows)
            : base(windows)
        {
        }

        public int Count
        {
            get { return ComObject.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public IApplication Parent
        {
            get { return new Application(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public IWindow this[object index]
        {
            get { return new Window(ComObject.Item(index)); }
        }

        public IWindow CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition, ref object docObj)
        {
            return new Window(ComObject.CreateToolWindow((Microsoft.Vbe.Interop.AddIn)addInInst.ComObject, progId, caption, guidPosition, ref docObj));
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

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Windows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(IWindows other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Windows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}