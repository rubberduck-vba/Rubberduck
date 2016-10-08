using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBProjects : SafeComWrapper<Microsoft.Vbe.Interop.VBProjects>, IEnumerable<VBProject>, IEquatable<VBProjects>
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

        public VBProject Add(ProjectType type)
        {
            return new VBProject(ComObject.Add((vbext_ProjectType)type));
        }

        public void Remove(VBProject project)
        {
            ComObject.Remove(project.ComObject);
        }

        public VBProject Open(string path)
        {
            return new VBProject(ComObject.Open(path));
        }

        public VBProject this[object index]
        {
            get { return new VBProject(ComObject.Item(index)); }
        }

        IEnumerator<VBProject> IEnumerable<VBProject>.GetEnumerator()
        {
            return new ComWrapperEnumerator<VBProject>(ComObject);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<VBProject>)this).GetEnumerator();
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

        public bool Equals(VBProjects other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.VBProjects>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}