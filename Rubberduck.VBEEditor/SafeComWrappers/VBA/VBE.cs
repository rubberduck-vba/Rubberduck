using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBE : SafeComWrapper<Microsoft.Vbe.Interop.VBE>, IVBE
    {
        public VBE(Microsoft.Vbe.Interop.VBE comObject)
            :base(comObject)
        {
        }

        public string Version
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Version; }
        }

        public ICodePane ActiveCodePane
        {
            get { return new CodePane(IsWrappingNullReference ? null : ComObject.ActiveCodePane); }
            set { ComObject.ActiveCodePane = (Microsoft.Vbe.Interop.CodePane)value.ComObject; }
        }

        public IVBProject ActiveVBProject
        {
            get { return new VBProject(IsWrappingNullReference ? null : ComObject.ActiveVBProject); }
            set { ComObject.ActiveVBProject = (Microsoft.Vbe.Interop.VBProject)value.ComObject; }
        }

        public IWindow ActiveWindow
        {
            get { return new Window(IsWrappingNullReference ? null : ComObject.ActiveWindow); }
        }

        public IAddIns AddIns
        {
            get { return new AddIns(IsWrappingNullReference ? null : ComObject.Addins); }
        }

        public ICodePanes CodePanes
        {
            get { return new CodePanes(IsWrappingNullReference ? null : ComObject.CodePanes); }
        }

        public ICommandBars CommandBars
        {
            get { return new CommandBars(IsWrappingNullReference ? null : ComObject.CommandBars); }
        }

        public IWindow MainWindow
        {
            get { return new Window(IsWrappingNullReference ? null : ComObject.MainWindow); }
        }

        public IVBComponent SelectedVBComponent
        {
            get { return new VBComponent(IsWrappingNullReference ? null : ComObject.SelectedVBComponent); }
        }

        public IVBProjects VBProjects
        {
            get { return new VBProjects(IsWrappingNullReference ? null : ComObject.VBProjects); }
        }

        public IWindows Windows
        {
            get { return new Windows(IsWrappingNullReference ? null : ComObject.Windows); }
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                VBProjects.Release();
                CodePanes.Release();
                CommandBars.Release();
                Windows.Release();
                AddIns.Release();
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.VBE> other)
        {
            return IsEqualIfNull(other) || (other != null && other.ComObject.Version == Version);
        }

        public bool Equals(IVBE other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBE>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }

        public bool IsInDesignMode
        {
            get { return VBProjects.All(project => project.Mode == EnvironmentMode.Design); }
        }

        public static void SetSelection(IVBProject vbProject, Selection selection, string name)
        {
            var components = vbProject.VBComponents;
            var component = components.SingleOrDefault(c => c.Name == name);
            if (component == null || component.IsWrappingNullReference)
            {
                return;
            }

            var module = component.CodeModule;
            var pane = module.CodePane;
            pane.SetSelection(selection);
        }
    }
}
