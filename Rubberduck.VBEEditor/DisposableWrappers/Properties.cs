using System.Collections;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Properties : SafeComWrapper<Microsoft.Vbe.Interop.Properties>, IEnumerable
    {
        public Properties(Microsoft.Vbe.Interop.Properties comObject) 
            : base(comObject)
        {
        }

        public Property Item(object index)
        {
            return new Property(InvokeResult(() => ComObject.Item(index)));
        }

        public Application Application { get { return new Application(InvokeResult(() => ComObject.Application)); } }
        /// <summary>
        /// Returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Parent { get { return InvokeResult(() => ComObject.Parent); } }
        public int Count { get { return InvokeResult(() => ComObject.Count); } }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public IEnumerator GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
        }
    }
}