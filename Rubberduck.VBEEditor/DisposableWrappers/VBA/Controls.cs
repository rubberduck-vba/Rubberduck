using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Controls : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Controls>, IEnumerable<Control>
    {
        public Controls(Microsoft.Vbe.Interop.Forms.Controls comObject) 
            : base(comObject)
        {
        }

        public int Count { get { return InvokeResult(() => ComObject.Count); } }

        public Control Item(object index)
        {
            return new Control(InvokeResult(() => (Microsoft.Vbe.Interop.Forms.Control)ComObject.Item(index)));
        }

        IEnumerator<Control> IEnumerable<Control>.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.Forms.Controls, Control>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new ComWrapperEnumerator<Microsoft.Vbe.Interop.Forms.Controls, Control>(ComObject);
        }
    }
}