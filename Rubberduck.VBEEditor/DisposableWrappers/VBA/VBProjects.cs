using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class VBProjects : SafeComWrapper<Microsoft.Vbe.Interop.VBProjects>, IEnumerable<VBProject>, IEquatable<VBProjects>
    {
        public VBProjects(Microsoft.Vbe.Interop.VBProjects comObject) 
            : base(comObject)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Count); }
        }

        public VBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : InvokeResult(() => ComObject.VBE)); }
        }

        public VBE Parent
        {
            get { return new VBE(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Parent)); }
        }

        public VBProject Add(ProjectType type)
        {
            return new VBProject(InvokeResult(() => ComObject.Add((vbext_ProjectType)type)));
        }

        public void Remove(VBProject project)
        {
            Invoke(() => ComObject.Remove(project.ComObject));
        }

        public VBProject Open(string path)
        {
            return new VBProject(InvokeResult(() => ComObject.Open(path)));
        }

        public VBProject Item(object index)
        {
            return new VBProject(InvokeResult(() => ComObject.Item(index)));
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
                    Item(i).Release();
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