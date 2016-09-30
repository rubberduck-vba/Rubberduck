using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class VBProject : WrapperBase<Microsoft.Vbe.Interop.VBProject>
    {
        public VBProject(Microsoft.Vbe.Interop.VBProject vbProject)
            :base(vbProject)
        {
        }

        public void SaveAs(string fileName)
        {
            ThrowIfDisposed();
            InvokeMember(path => Item.SaveAs(path), fileName);
        }

        public void MakeCompiledFile()
        {
            ThrowIfDisposed();
            InvokeMember(() => Item.MakeCompiledFile());
        }

        public Application Application
        {
            get
            {
                ThrowIfDisposed();
                return new Application(InvokeMemberValue(() => Item.Application));
            }
        }

        public Application Parent
        {
            get
            {
                ThrowIfDisposed();
                return new Application(InvokeMemberValue(() => Item.Parent));
            }
        }

        public string HelpFile
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.HelpFile);
            }
            set
            {
                ThrowIfDisposed();
                Item.HelpFile = value;
            }
        }

        public int HelpContextID
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.HelpContextID);
            }
            set
            {
                ThrowIfDisposed();
                Item.HelpContextID = value;
            }
        }

        public string Description {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.Description);
            }
            set
            {
                ThrowIfDisposed();
                Item.Description = value;
            } 
        }

        public References References
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public string Name
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.Name);
            }
            set
            {
                ThrowIfDisposed();
                Item.Name = value;
            }
        }

        public EnvironmentMode Mode
        {
            get
            {
                ThrowIfDisposed();
                return (EnvironmentMode)InvokeMemberValue(() => Item.Mode);
            }
        }

        public VBProjects Collection
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public ProjectProtection Protection
        {
            get
            {
                ThrowIfDisposed();
                return (ProjectProtection)InvokeMemberValue(() => Item.Protection);
            }
        }

        public bool Saved
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.Saved);
            }
        }

        public VBComponents VBComponents
        {
            get
            {
                ThrowIfDisposed();
                throw new NotImplementedException();
            }
        }

        public ProjectType Type
        {
            get
            {
                ThrowIfDisposed();
                return (ProjectType)InvokeMemberValue(() => Item.Type);
            }
        }

        public string FileName
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.FileName);
            }
        }

        public string BuildFileName
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.BuildFileName);
            }
        }

        public VBE VBE
        {
            get
            {
                ThrowIfDisposed();
                return new VBE(InvokeMemberValue(() => Item.VBE));
            }
        }
    }
}