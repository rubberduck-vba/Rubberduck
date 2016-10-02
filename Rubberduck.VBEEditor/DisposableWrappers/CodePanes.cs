using System.Collections;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class CodePanes : SafeComWrapper<Microsoft.Vbe.Interop.CodePanes>, IEnumerable
    {
        public CodePanes(Microsoft.Vbe.Interop.CodePanes comObject) 
            : base(comObject)
        {
        }

        public CodePane Item(object index)
        {
            return new CodePane(InvokeResult(() => ComObject.Item(index)));
        }

        public VBE Parent { get { return new VBE(InvokeResult(() => ComObject.Parent)); } }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
        public int Count { get { return InvokeResult(() => ComObject.Count); } }
        public CodePane Current 
        { 
            get{ return new CodePane(InvokeResult(() => ComObject.Current)); }
            set{ Invoke(() => ComObject.Current = value.ComObject);}
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return InvokeResult(() => ComObject.GetEnumerator());
        }
    }
}