
namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class AddIn : SafeComWrapper<Microsoft.Vbe.Interop.AddIn>
    {
        public AddIn(Microsoft.Vbe.Interop.AddIn comObject) 
            : base(comObject)
        {
        }

        public string Description
        {
            get { return InvokeResult(() => ComObject.Description); }
            set { Invoke(() => ComObject.Description = value); }
        }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
        public AddIns Collection { get { return new AddIns(InvokeResult(() => ComObject.Collection)); } }
        public string ProgId { get { return InvokeResult(() => ComObject.ProgId); } }
        public string Guid { get { return InvokeResult(() => ComObject.Guid); } }

        public bool Connect
        {
            get { return InvokeResult(() => ComObject.Connect); }
            set { Invoke(() => ComObject.Connect = value); }
        }

        public object Object // definitely leaks a COM object
        {
            get { return InvokeResult(() => ComObject.Object); }
            set { Invoke(() => ComObject.Object = value); }
        }
    }
}