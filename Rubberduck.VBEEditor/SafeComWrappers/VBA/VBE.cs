using System;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBE : SafeComWrapper<Microsoft.Vbe.Interop.VBE>, IEquatable<VBE>
    {
        public VBE(Microsoft.Vbe.Interop.VBE comObject)
            :base(comObject)
        {
        }

        public CodePane ActiveCodePane
        {
            get { return new CodePane(IsWrappingNullReference ? null : InvokeResult(() => ComObject.ActiveCodePane)); }
            set { Invoke(() => ComObject.ActiveCodePane = value.ComObject); }
        }

        public VBProject ActiveVBProject
        {
            get { return new VBProject(IsWrappingNullReference ? null : InvokeResult(() => ComObject.ActiveVBProject)); }
            set { Invoke(() => ComObject.ActiveVBProject = value.ComObject); }
        }

        public Window ActiveWindow
        {
            get { return new Window(IsWrappingNullReference ? null : InvokeResult(() => ComObject.ActiveWindow)); }
        }

        public AddIns AddIns
        {
            get { return new AddIns(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Addins)); }
        }

        public CodePanes CodePanes
        {
            get { return new CodePanes(IsWrappingNullReference ? null : InvokeResult(() => ComObject.CodePanes)); }
        }

        public CommandBars CommandBars
        {
            get { return new CommandBars(IsWrappingNullReference ? null : InvokeResult(() => ComObject.CommandBars)); }
        }

        public Window MainWindow
        {
            get { return new Window(IsWrappingNullReference ? null : InvokeResult(() => ComObject.MainWindow)); }
        }

        public VBComponent SelectedVBComponent
        {
            get { return new VBComponent(IsWrappingNullReference ? null : InvokeResult(() => ComObject.SelectedVBComponent)); }
        }

        public VBProjects VBProjects
        {
            get { return new VBProjects(IsWrappingNullReference ? null : InvokeResult(() => ComObject.VBProjects)); }
        }

        public string Version
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Version); }
        }

        public Windows Windows
        {
            get { return new Windows(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Windows)); }
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

        public bool Equals(VBE other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBE>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }

        public bool IsInDesignMode()
        {
            return VBProjects.All(project => project.Mode == EnvironmentMode.Design);
        }

        public static void SetSelection(VBProject vbProject, Selection selection, string name)
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
