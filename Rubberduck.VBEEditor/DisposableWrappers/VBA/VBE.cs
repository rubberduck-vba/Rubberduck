namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class VBE : SafeComWrapper<Microsoft.Vbe.Interop.VBE>
    {
        public VBE(Microsoft.Vbe.Interop.VBE comObject)
            :base(comObject)
        {
        }

        public CodePane ActiveCodePane
        {
            get { return new CodePane(InvokeResult(() => ComObject.ActiveCodePane)); }
            set { Invoke(() => ComObject.ActiveCodePane = value.ComObject); }
        }

        public VBProject ActiveVBProject
        {
            get { return new VBProject(InvokeResult(() => ComObject.ActiveVBProject)); }
            set { Invoke(() => ComObject.ActiveVBProject = value.ComObject); }
        }

        public Window ActiveWindow { get { return new Window(InvokeResult(() => ComObject.ActiveWindow)); } }

        public AddIns AddIns { get { return new AddIns(InvokeResult(() => ComObject.Addins)); } }

        public CodePanes CodePanes { get { return new CodePanes(InvokeResult(() => ComObject.CodePanes)); } }

        /// <summary>
        /// Returns an unwrapped COM object; remember to call Marshal.ReleaseComObject on the returned object.
        /// </summary>
        public Microsoft.Office.Core.CommandBars CommandBars { get { return InvokeResult(() => ComObject.CommandBars); } }

        public Window MainWindow { get { return new Window(InvokeResult(() => ComObject.MainWindow)); } }

        public VBComponent SelectedVBComponent { get { return new VBComponent(InvokeResult(() => ComObject.SelectedVBComponent)); } }

        public VBProjects VBProjects { get { return new VBProjects(InvokeResult(() => ComObject.VBProjects)); } }

        public string Version { get { return InvokeResult(() => ComObject.Version); } }

        public Windows Windows { get { return new Windows(InvokeResult(() => ComObject.Windows)); } }
    }
}
