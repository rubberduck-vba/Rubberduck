namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Control : SafeComWrapper<Microsoft.Vbe.Interop.Forms.Control>
    {
        public Control(Microsoft.Vbe.Interop.Forms.Control comObject) 
            : base(comObject)
        {
        }

        public string Name
        {
            get { return InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }
    }
}