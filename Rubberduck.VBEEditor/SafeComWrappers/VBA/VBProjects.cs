using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBProjects : SafeComWrapper<VB.VBProjects>, IVBProjects
    {
        private static readonly Guid VBProjectsEventsGuid = new Guid("0002E103-0000-0000-C000-000000000046");

        private enum ProjectEventDispId
        {
            ItemAdded = 1,
            ItemRemoved = 2,
            ItemRenamed = 3,
            ItemActivated = 4
        }

        public VBProjects(VB.VBProjects target) : base(target)
        {
            AttachEvents();
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
                DetatchEvents();
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

        #region Events

        private bool _eventsAttached;
        private void AttachEvents()
        {
            if (!_eventsAttached && !IsWrappingNullReference)
            {
                _projectAdded = OnProjectAdded;
                _projectRemoved = OnProjectRemoved;
                _projectRenamed = OnProjectRenamed;
                _projectActivated = OnProjectActivated;
                ComEventsHelper.Combine(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemAdded, _projectAdded);
                ComEventsHelper.Combine(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRemoved, _projectRemoved);
                ComEventsHelper.Combine(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRenamed, _projectRenamed);
                ComEventsHelper.Combine(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemActivated, _projectActivated);
            }
        }

        private void DetatchEvents()
        {
            if (!_eventsAttached && !IsWrappingNullReference)
            {
                ComEventsHelper.Remove(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemAdded, _projectAdded);
                ComEventsHelper.Remove(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRemoved, _projectRemoved);
                ComEventsHelper.Remove(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRenamed, _projectRenamed);
                ComEventsHelper.Remove(Target, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemActivated, _projectActivated);
                _eventsAttached = false;
            }
        }

        public event EventHandler<ProjectEventArgs> ProjectAdded;
        private delegate void ItemAddedDelegate(VB.VBProject vbProject);
        private ItemAddedDelegate _projectAdded;
        private void OnProjectAdded(VB.VBProject vbProject)
        {
            if (VBE.IsInDesignMode) OnDispatch(ProjectAdded, vbProject, true);
        }

        public event EventHandler<ProjectEventArgs> ProjectRemoved;
        private delegate void ItemRemovedDelegate(VB.VBProject vbProject);
        private ItemRemovedDelegate _projectRemoved;
        private void OnProjectRemoved(VB.VBProject vbProject)
        {
            if (VBE.IsInDesignMode) OnDispatch(ProjectRemoved, vbProject);
        }

        public event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;
        private delegate void ItemRenamedDelegate(VB.VBProject vbProject, string oldName);
        private ItemRenamedDelegate _projectRenamed;
        private void OnProjectRenamed(VB.VBProject vbProject, string oldName)
        {
            if (!VBE.IsInDesignMode) { return; }

            var project = new VBProject(vbProject);
            var projectId = project.ProjectId;

            var handler = ProjectRenamed;
            if (handler != null)
            {
                handler(this, new ProjectRenamedEventArgs(projectId, project, oldName));
            }
        }

        public event EventHandler<ProjectEventArgs> ProjectActivated;
        private delegate void ItemActivatedDelegate(VB.VBProject vbProject);
        private ItemActivatedDelegate _projectActivated;
        private void OnProjectActivated(VB.VBProject vbProject)
        {
            if (VBE.IsInDesignMode) OnDispatch(ProjectActivated, vbProject);
        }

        private void OnDispatch(EventHandler<ProjectEventArgs> dispatched, VB.VBProject vbProject, bool assignId = false)
        {
            var handler = dispatched;
            if (handler != null)
            {
                var project = new VBProject(vbProject);
                if (assignId)
                {
                    project.AssignProjectId();
                }
                var projectId = project.ProjectId;
                handler.Invoke(this, new ProjectEventArgs(projectId, project));
            }
        }

        #endregion
    }
}