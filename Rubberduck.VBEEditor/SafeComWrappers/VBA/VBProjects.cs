using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBProjects : SafeComWrapper<VB.VBProjects>, IVBProjects
    {
        private readonly IVBProjectsEventsSink _sink;

        public VBProjects(VB.VBProjects target) 
            : base(target)
        {
            _sink = new VBProjectsEventsSink();
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
            return new VBProject(IsWrappingNullReference ? null : Target.Add((VB.vbext_ProjectType)type));
        }

        public void Remove(IVBProject project)
        {
            if (IsWrappingNullReference) return;
            Target.Remove((VB.VBProject) project.Target);
        }

        public IVBProjectsEventsSink Events { get { return _sink; } }

        public IConnectionPoint ConnectionPoint
        {
            get
            {
                try
                {
                    var connectionPointContainer = (IConnectionPointContainer)Target;
                    var interfaceId = typeof(VB._dispVBProjectsEvents).GUID;
                    IConnectionPoint result;
                    connectionPointContainer.FindConnectionPoint(ref interfaceId, out result);
                    return result;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        public IVBProject Open(string path)
        {
            return new VBProject(IsWrappingNullReference ? null : Target.Open(path));
        }

        public IVBProject this[object index]
        {
            get { return new VBProject(IsWrappingNullReference ? null : Target.Item(index)); }
        }

        IEnumerator<IVBProject> IEnumerable<IVBProject>.GetEnumerator()
        {
            return IsWrappingNullReference
                ? new ComWrapperEnumerator<IVBProject>(null, o => new VBProject(null))
                : new ComWrapperEnumerator<IVBProject>(Target, o => new VBProject((VB.VBProject) o));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IVBProject>) this).GetEnumerator();
        }

        public override void Release(bool final = false)
        {
            if (!IsWrappingNullReference)
            {
                for (var i = 1; i <= Count; i++)
                {
                    this[i].Release();
                }
                base.Release(final);
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