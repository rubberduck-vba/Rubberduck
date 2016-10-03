using System.Collections;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Windows : SafeComWrapper<Microsoft.Vbe.Interop.Windows>, IEnumerable
    {
        public Windows(Microsoft.Vbe.Interop.Windows windows)
            : base(windows)
        {
        }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public Application Parent { get { return new Application(InvokeResult(() => ComObject.Parent)); } }

        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        public Window CreateToolWindow(AddIn addInInst, string progId, string caption, string guidPosition, ref object docObj)
        {
            ThrowIfDisposed();
            try
            {
                return new Window(ComObject.CreateToolWindow(addInInst.ComObject, progId, caption, guidPosition, ref docObj));
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        public Window Item(object index)
        {
            return new Window(InvokeResult(() => ComObject.Item(index)));
        }

        public IEnumerator GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
        }
    }
}