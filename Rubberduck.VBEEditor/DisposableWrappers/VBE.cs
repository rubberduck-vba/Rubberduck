using System;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class VBE : WrapperBase<Microsoft.Vbe.Interop.VBE>
    {
        public VBE(Microsoft.Vbe.Interop.VBE vbe)
            :base(vbe)
        {
        }

        public CodePane ActiveCodePane
        {
            get
            {
                ThrowIfDisposed();
                return new CodePane(InvokeMemberValue(() => Item.ActiveCodePane));
            }
            set
            {
                ThrowIfDisposed();
                Item.ActiveCodePane = value.Item;
            }
        }

        public VBProject ActiveVBProject
        {
            get
            {
                ThrowIfDisposed();
                return new VBProject(InvokeMemberValue(() => Item.ActiveVBProject));
            }
            set
            {
                ThrowIfDisposed();
                Item.ActiveVBProject = value.Item;
            }
        }

        public Window ActiveWindow
        {
            get
            {
                ThrowIfDisposed();
                return new Window(InvokeMemberValue(() => Item.ActiveWindow));
            }
        }

        public Microsoft.Vbe.Interop.Addins Addins
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public Microsoft.Vbe.Interop.CodePanes CodePanes
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public Microsoft.Office.Core.CommandBars CommandBars
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public Microsoft.Vbe.Interop.Events Events
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public Window MainWindow
        {
            get
            {
                ThrowIfDisposed();
                return new Window(InvokeMemberValue(() => Item.MainWindow));
            }
        }

        public Microsoft.Vbe.Interop.VBComponent SelectedVBComponent
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public Microsoft.Vbe.Interop.VBProjects VBProjects
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public string Version
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.Version);
            }
        }

        public Windows Windows
        {
            get
            {
                ThrowIfDisposed();
                return new Windows(InvokeMemberValue(() => Item.Windows));
            }
        }
    }
}
