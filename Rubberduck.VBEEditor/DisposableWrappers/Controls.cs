using System.Collections;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Controls : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>, IEnumerable
    {
        public Controls(Microsoft.Vbe.Interop.Forms.Controls comObject) 
            : base(comObject)
        {
        }

        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        public Control Item(object index)
        {
            return new Control(InvokeResult(() => ComObject.Item(index)));
        }

        public IEnumerator GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
        }
    }
}