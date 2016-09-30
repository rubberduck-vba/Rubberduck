using System.Collections;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Windows : WrapperBase<Microsoft.Vbe.Interop.Windows>, IEnumerable
    {
        public Windows(Microsoft.Vbe.Interop.Windows windows)
            : base(windows)
        {
        }

        public VBE VBE
        {
            get
            {
                ThrowIfDisposed();
                return new VBE(InvokeMemberValue(() => base.Item.VBE));
            }
        }

        public Application Parent
        {
            get
            {
                ThrowIfDisposed();
                return new Application(InvokeMemberValue(() => base.Item.Parent));
            }
        }

        public int Count
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => base.Item.Count);
            }
        }

        public Window CreateToolWindow(AddIn addInInst, string progId, string caption, string guidPosition, ref object docObj)
        {
            ThrowIfDisposed();
            return new Window(base.Item.CreateToolWindow(addInInst, progId, caption, guidPosition, ref docObj));
        }

        public Window Item(object index)
        {
            ThrowIfDisposed();
            return new Window(InvokeMemberValue(i => base.Item.Item(i), index));
        }

        public IEnumerator GetEnumerator()
        {
            ThrowIfDisposed();
            return InvokeMemberValue(() => base.Item.GetEnumerator());
        }
    }
}