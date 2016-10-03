using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers
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
            return new ComWrapperEnumerator<Controls, Control>(this);
        }

        public IEnumerator GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
        }
    }
}