using System.Collections;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class AddIns : SafeComWrapper<Microsoft.Vbe.Interop.Addins>, IEnumerable
    {
        public AddIns(Microsoft.Vbe.Interop.Addins comObject) : 
            base(comObject)
        {
        }

        public AddIn Item(object index)
        {
            return new AddIn(InvokeResult(() => ComObject.Item(index)));
        }

        public void Update()
        {
            Invoke(() => ComObject.Update());
        }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        /// <summary>
        /// Getter returns an unwrapped COM object representing the host application; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Parent { get { return InvokeResult(() => ComObject.Parent); } } 
        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        public IEnumerator GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
        }
    }
}