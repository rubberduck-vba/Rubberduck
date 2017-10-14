using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;
using VBAIA = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.VBA
{
    public class VBProjects : VBProjects<VBAIA.VBProjects>
    {        
        public VBProjects(VBAIA.VBProjects target) : base(target, VBType.VBA)
        {   
        }

        public override int Count => IsWrappingNullReference ? 0 : Target.Count;

        public override IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public override IVBE Parent => new VBE(IsWrappingNullReference ? null : Target.Parent);

        public override IVBProject Add(ProjectType type)
        {
            return new VBProject(IsWrappingNullReference ? null : Target.Add((VBAIA.vbext_ProjectType)type));
        }

        public override void Remove(IVBProject project)
        {
            if (IsWrappingNullReference) return;
            Target.Remove((VBAIA.VBProject) project.Target);
        }

        public override IVBProject Open(string path)
        {
            return new VBProject(IsWrappingNullReference ? null : Target.Open(path));
        }

        public override IVBProject this[object index] => new VBProject(IsWrappingNullReference ? null : Target.Item(index));
        

        public override IEnumerator<IVBProject> GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<IVBProject>(null, o => new VBProject(null))
                : new ComWrapperEnumerator<IVBProject>(Target, o => new VBProject((VBAIA.VBProject)o));
        }

        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        for (var i = 1; i <= Count; i++)
        //        {
        //            this[i].Release();
        //        }
        //        base.Release(final);
        //    }
        //}

        public override bool Equals(IVBProjects other)
        {
            return ((other == null || other.IsWrappingNullReference) && IsWrappingNullReference)
                   || (other != null && !IsWrappingNullReference && ReferenceEquals(other.Target, Target));
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 
                : HashCode.Compute(Target);
        }
    }
}