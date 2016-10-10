using System;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBE : SafeComWrapper<VB.VBE>, IVBE
    {
        public VBE(VB.VBE target)
            :base(target)
        {
        }

        public string Version
        {
            get { return IsWrappingNullReference ? string.Empty : Target.get_Version(); }
        }

        public ICodePane ActiveCodePane
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public IVBProject ActiveVBProject
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public IWindow ActiveWindow
        {
            get { throw new NotImplementedException(); }
        }

        public IAddIns AddIns
        {
            get { throw new NotImplementedException(); }
        }

        public ICodePanes CodePanes
        {
            get { throw new NotImplementedException(); }
        }

        public ICommandBars CommandBars
        {
            get { throw new NotImplementedException(); }
        }

        public IWindow MainWindow
        {
            get { throw new NotImplementedException(); }
        }

        public IVBComponent SelectedVBComponent
        {
            get { throw new NotImplementedException(); }
        }

        public IVBProjects VBProjects
        {
            get { return new VBProjects(IsWrappingNullReference ? null : Target.get_VBProjects()); }
        }

        public IWindows Windows
        {
            get { throw new NotImplementedException(); }
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
                Marshal.ReleaseComObject(Target);
            }
        }

        public override bool Equals(ISafeComWrapper<VB.VBE> other)
        {
            return IsEqualIfNull(other) || (other != null && other.Target.get_Version() == Version);
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
            pane.SetSelection(selection);
        }
    }
}
