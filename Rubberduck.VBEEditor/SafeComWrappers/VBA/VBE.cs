using System;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBE : SafeComWrapper<VB.VBE>, IVBE
    {
        public VBE(VB.VBE target)
            : base(target)
        {
        }

        public string Version
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Version; }
        }

        public ICodePane ActiveCodePane
        {
            get { return new CodePane(IsWrappingNullReference ? null : Target.ActiveCodePane); }
            set { if (!IsWrappingNullReference) Target.ActiveCodePane = (VB.CodePane)value.Target; }
        }

        public IVBProject ActiveVBProject
        {
            get { return new VBProject(IsWrappingNullReference ? null : Target.ActiveVBProject); }
            set { if (!IsWrappingNullReference) Target.ActiveVBProject = (VB.VBProject)value.Target; }
        }

        public IWindow ActiveWindow
        {
            get { return new Window(IsWrappingNullReference ? null : Target.ActiveWindow); }
        }

        public IAddIns AddIns
        {
            get { return new AddIns(IsWrappingNullReference ? null : Target.Addins); }
        }

        public ICodePanes CodePanes
        {
            get { return new CodePanes(IsWrappingNullReference ? null : Target.CodePanes); }
        }

        public ICommandBars CommandBars
        {
            get { return new CommandBars(IsWrappingNullReference ? null : Target.CommandBars); }
        }

        public IWindow MainWindow
        {
            get
            {
                try
                {
                    return new Window(IsWrappingNullReference ? null : Target.MainWindow);
                }
                catch (InvalidComObjectException)
                {
                    return null;
                }
            }
        }

        public IVBComponent SelectedVBComponent
        {
            get { return new VBComponent(IsWrappingNullReference ? null : Target.SelectedVBComponent); }
        }

        public IVBProjects VBProjects
        {
            get { return new VBProjects(IsWrappingNullReference ? null : Target.VBProjects); }
        }

        public IWindows Windows
        {
            get { return new Windows(IsWrappingNullReference ? null : Target.Windows); }
        }

        public Guid EventsInterfaceId { get { throw new NotImplementedException(); } }

        public override void Release(bool final = false)
        {
            if (!IsWrappingNullReference)
            {
                VBProjects.Release();
                CodePanes.Release();
                //CommandBars.Release();
                Windows.Release();
                AddIns.Release();
                base.Release(final);
            }
        }

        public override bool Equals(ISafeComWrapper<VB.VBE> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.Version == Version);
        }

        public bool Equals(IVBE other)
        {
            return Equals(other as SafeComWrapper<VB.VBE>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
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
            pane.Selection = selection;
        }
    }
}
