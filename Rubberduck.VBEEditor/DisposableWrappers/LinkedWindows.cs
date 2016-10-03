using System.Collections;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class LinkedWindows : SafeComWrapper<Microsoft.Vbe.Interop.LinkedWindows>, IEnumerable
    {
        public LinkedWindows(Microsoft.Vbe.Interop.LinkedWindows linkedWindows)
            : base(linkedWindows)
        {
        }

        public Window Item(object index) { return new Window(InvokeResult(() => ComObject.Item(index))); }

        public void Remove(Window window) { Invoke(() => ComObject.Remove(window.ComObject)); }

        public void Add(Window window) { Invoke(() => ComObject.Add(window.ComObject)); }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public Window Parent { get { return new Window(InvokeResult(() => ComObject.Parent)); } }

        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        public IEnumerator GetEnumerator() { return InvokeResult(() => ComObject.GetEnumerator()); }
    }
}