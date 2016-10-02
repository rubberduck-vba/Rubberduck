using System;

namespace Rubberduck.VBEditor.DisposableWrappers
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
            set { Invoke(o => ComObject.ActiveCodePane = o, value.ComObject); }
        }

        public VBProject ActiveVBProject
        {
            get { return new VBProject(InvokeResult(() => ComObject.ActiveVBProject)); }
            set { Invoke(o => ComObject.ActiveVBProject = o, value.ComObject); }
        }

        public Window ActiveWindow { get { return new Window(InvokeResult(() => ComObject.ActiveWindow)); } }

        public AddIns AddIns { get { return new AddIns(InvokeResult(() => ComObject.Addins)); } }

        public CodePanes CodePanes { get { return new CodePanes(InvokeResult(() => ComObject.CodePanes)); } }

        public Microsoft.Office.Core.CommandBars CommandBars
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public Window MainWindow { get { return new Window(InvokeResult(() => ComObject.MainWindow)); } }

        public Microsoft.Vbe.Interop.VBComponent SelectedVBComponent
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public Microsoft.Vbe.Interop.VBProjects VBProjects
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public string Version { get { return InvokeResult(() => ComObject.Version); } }

        public Windows Windows { get { return new Windows(InvokeResult(() => ComObject.Windows)); } }
    }
}
