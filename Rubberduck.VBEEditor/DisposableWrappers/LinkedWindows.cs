using System.Collections;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class LinkedWindows : WrapperBase<Microsoft.Vbe.Interop.LinkedWindows>, IEnumerable
    {
        public LinkedWindows(Microsoft.Vbe.Interop.LinkedWindows linkedWindows)
            : base(linkedWindows)
        {
        }

        public Window Item(object index)
        {
            ThrowIfDisposed();
            return new Window(InvokeMemberValue(item => base.Item.Item(item), index));
        }

        public void Remove(Window window)
        {
            ThrowIfDisposed();
            InvokeMember(item => base.Item.Remove(item), window);
        }

        public void Add(Window window)
        {
            ThrowIfDisposed();
            InvokeMember(item => base.Item.Add(item), window);
        }

        public VBE VBE
        {
            get
            {
                ThrowIfDisposed();
                return new VBE(InvokeMemberValue(() => base.Item.VBE));
            }
        }

        public Window Parent
        {
            get
            {
                ThrowIfDisposed();
                return new Window(InvokeMemberValue(() => base.Item.Parent));
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

        public IEnumerator GetEnumerator()
        {
            ThrowIfDisposed();
            return InvokeMemberValue(() => base.Item.GetEnumerator());
        }
    }
}