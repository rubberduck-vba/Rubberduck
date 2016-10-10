using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VBProjects : SafeComWrapper<VB.VBProjects>, IVBProjects
    {
        public VBProjects(VB.VBProjects target) 
            : base(target)
        {
        }

        public int Count
        {
            get { return IsWrappingNullReference ? 0 : Target.Count; }
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IVBE Parent
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.Parent); }
        }

        public IVBProject Add(ProjectType type)
        {
            return new VBProject(Target.Add((VB.vbext_ProjectType)type));
        }

        public IVBProject Open(string path)
        {
            throw new NotImplementedException();
        }

        public void Remove(IVBProject project)
        {
            Target.Remove((VB.VBProject)project.Target);
        }

        public IVBProject this[object index]
        {
            get { return new VBProject(Target.Item(index)); }
        }

        IEnumerator<IVBProject> IEnumerable<IVBProject>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IVBProject>(Target, o => new VBProject((VB.VBProject)o));
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
                Marshal.ReleaseComObject(Target);
            }
        }

        public override bool Equals(ISafeComWrapper<VB.VBProjects> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(IVBProjects other)
        {
            return Equals(other as SafeComWrapper<VB.VBProjects>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 
                : HashCode.Compute(Target);
        }
    }
}