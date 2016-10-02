namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class VBComponent : SafeComWrapper<Microsoft.Vbe.Interop.VBComponent>
    {
        public VBComponent(Microsoft.Vbe.Interop.VBComponent comObject) 
            : base(comObject)
        {
        }

        public Window DesignerWindow()
        {
            return new Window(InvokeResult(() => ComObject.DesignerWindow()));
        }

        public void Activate()
        {
            Invoke(() => ComObject.Activate());
        }

        public bool IsSaved { get { return InvokeResult(() => ComObject.Saved); } }

        public string Name
        {
            get { return InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        public void Export(string path)
        {
            Invoke(() => ComObject.Export(path));
        }

        /// <summary>
        /// Returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public object Designer { get { return InvokeResult(() => ComObject.Designer); } }

        public CodeModule CodeModule { get { return new CodeModule(InvokeResult(() => ComObject.CodeModule)); } }
        public ComponentType Type { get { return (ComponentType)InvokeResult(() => ComObject.Type); } }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }
        public VBComponents Collection { get { return new VBComponents(InvokeResult(() => ComObject.Collection)); } }
        public bool HasOpenDesigner { get { return InvokeResult(() => ComObject.HasOpenDesigner); } }
        public Properties Properties { get { return new Properties(InvokeResult(() => ComObject.Properties)); } }
        public string DesignerId { get { return InvokeResult(() => ComObject.DesignerID); } }
    }
}