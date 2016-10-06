using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBProjects : SafeComWrapper<Microsoft.Vbe.Interop.VBProjects>, IVBProjects
    {
        public VBProjects(Microsoft.Vbe.Interop.VBProjects comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }

        public IVBE Parent
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.Parent); }
        }

        public IVBProject Add(ProjectType type)
        {
            return new VBProject(ComObject.Add((vbext_ProjectType)type));
        }

        public void Remove(IVBProject project)
        {
            ComObject.Remove((Microsoft.Vbe.Interop.VBProject)project.ComObject);
        }

        public IVBProject Open(string path)
        {
            return new VBProject(ComObject.Open(path));
        }

        public IVBProject this[object index]
        {
            get { return new VBProject(ComObject.Item(index)); }
        }

        IEnumerator<IVBProject> IEnumerable<IVBProject>.GetEnumerator()
        {
            return new ComWrapperEnumerator<VBProject>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<IVBProject>)this).GetEnumerator();
        }

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.VBProjects> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.ComObject, ComObject));
        }

        public bool Equals(IVBProjects other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBProjects>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}