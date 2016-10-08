using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

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
            get { return Target.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IApplication Parent
        {
            get { return new Application(IsWrappingNullReference ? null : Target.Parent); }
        }

        public IWindow this[object index]
        {
            get { return new Window(Target.Item(index)); }
        }

        public IWindow CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition, ref object docObj)
        {
            return new Window(Target.CreateToolWindow((Microsoft.Vbe.Interop.AddIn)addInInst.Target, progId, caption, guidPosition, ref docObj));
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

        public override bool Equals(ISafeComWrapper<Microsoft.Vbe.Interop.Windows> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IWindows other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Windows>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}