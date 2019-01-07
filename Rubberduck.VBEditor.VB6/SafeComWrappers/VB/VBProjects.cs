using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public sealed class VBProjects : SafeEventedComWrapper<VB.VBProjects, VB._dispVBProjectsEvents>, IVBProjects, VB._dispVBProjectsEvents
    {
        public VBProjects(VB.VBProjects target, bool rewrapping = false)
        :base(target, rewrapping)
        {
        }

        public int Count => IsWrappingNullReference ? 0 : Target.Count;

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IVBE Parent => new VBE(IsWrappingNullReference ? null : Target.Parent);

        public IVBProject Add(ProjectType type)
        {
            return new VBProject(IsWrappingNullReference ? null : Target.Add((VB.vbext_ProjectType)type));
        }

        public void Remove(IVBProject project)
        {
            if (IsWrappingNullReference)
            {
                return;
            }
            Target.Remove((VB.VBProject)project.Target);
        }

        public IVBProject Open(string path) => new VBProject(IsWrappingNullReference? null : Target.AddFromFile(path).Item(1));

        public IVBProject StartProject
        {
            get => new VBProject(IsWrappingNullReference ? null : Target.StartProject);
            set => Target.StartProject = (VB.VBProject)value.Target;
        }

        public IVBProject this[object index] => new VBProject(IsWrappingNullReference ? null : Target.Item(index));

        IEnumerator<IVBProject> IEnumerable<IVBProject>.GetEnumerator()
        {
            return new ComWrapperEnumerator<IVBProject>(Target, comObject => new VBProject((VB.VBProject) comObject));
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return IsWrappingNullReference
                ? (IEnumerator) new List<IEnumerable>().GetEnumerator()
                : ((IEnumerable<IVBProject>) this).GetEnumerator();
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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);

        #region Events

        public event EventHandler<ProjectEventArgs> ProjectAdded;
        void VB._dispVBProjectsEvents.ItemAdded([MarshalAs(UnmanagedType.Interface), In] VB.VBProject VBProject)
        {
            OnDispatch(ProjectAdded, VBProject, true);
        }

        public event EventHandler<ProjectEventArgs> ProjectRemoved;
        void VB._dispVBProjectsEvents.ItemRemoved([MarshalAs(UnmanagedType.Interface), In] VB.VBProject VBProject)
        {
            OnDispatch(ProjectRemoved, VBProject);
        }

        public event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;
        void VB._dispVBProjectsEvents.ItemRenamed([MarshalAs(UnmanagedType.Interface), In] VB.VBProject VBProject,
            [MarshalAs(UnmanagedType.BStr), In] string OldName)
        {
            using (var project = new VBProject(VBProject))
            {
                if (!IsInDesignMode())
                {
                    return;
                }

                var projectId = project.ProjectId;

                if (projectId == null)
                {
                    return;
                }

                var handler = ProjectRenamed;
                handler?.Invoke(this, new ProjectRenamedEventArgs(projectId, project.Name, OldName));
            }
        }

        public event EventHandler<ProjectEventArgs> ProjectActivated;
        void VB._dispVBProjectsEvents.ItemActivated([MarshalAs(UnmanagedType.Interface), In] VB.VBProject VBProject)
        {
            OnDispatch(ProjectActivated, VBProject);
        }

        private void OnDispatch(EventHandler<ProjectEventArgs> dispatched, VB.VBProject vbProject, bool assignId = false)
        {
            using (var project = new VBProject(vbProject))
            {
                var handler = dispatched;
                if (handler == null || !IsInDesignMode())
                {
                    return;
                }

                if (assignId)
                {
                    project.AssignProjectId();
                }

                var projectId = project.ProjectId;

                if (projectId == null)
                {
                    return;
                }

                handler.Invoke(project, new ProjectEventArgs(projectId, project.Name));
            }
        }

        private bool IsInDesignMode()
        {
            foreach (var project in this)
                using(project)
                {
                    if (project.Mode != EnvironmentMode.Design)
                    {
                        return false;
                    }
                }
            return true;
        }

        #endregion
    }
}