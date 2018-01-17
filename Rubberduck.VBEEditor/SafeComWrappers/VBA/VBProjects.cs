using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBProjects : SafeComWrapper<VB.VBProjects>, IVBProjects
    {
        private static readonly Guid VBProjectsEventsGuid = new Guid("0002E103-0000-0000-C000-000000000046");
        private static VB.VBProjects _projects;
        private enum ProjectEventDispId
        {
            ItemAdded = 1,
            ItemRemoved = 2,
            ItemRenamed = 3,
            ItemActivated = 4
        }

        public VBProjects(VB.VBProjects target, bool rewrapping = false) 
        :base(target, rewrapping)
        {
            if (_projects == null)
            {
                _projects = target;
                AttachEvents();
            }          
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
            Target.Remove((VB.VBProject) project.Target);
        }

        public IVBProject Open(string path)
        {
            return new VBProject(IsWrappingNullReference ? null : Target.Open(path));
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

        #region Events

        private static void AttachEvents()
        {
            if (_projects != null)
            {
                _projectAdded = OnProjectAdded;
                _projectRemoved = OnProjectRemoved;
                _projectRenamed = OnProjectRenamed;
                _projectActivated = OnProjectActivated;
                ComEventsHelper.Combine(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemAdded, _projectAdded);
                ComEventsHelper.Combine(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRemoved, _projectRemoved);
                ComEventsHelper.Combine(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRenamed, _projectRenamed);
                ComEventsHelper.Combine(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemActivated, _projectActivated);
            }
        }

        internal static void DetatchEvents()
        {
            if (_projects != null)
            {
                ComEventsHelper.Remove(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemAdded, _projectAdded);
                ComEventsHelper.Remove(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRemoved, _projectRemoved);
                ComEventsHelper.Remove(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemRenamed, _projectRenamed);
                ComEventsHelper.Remove(_projects, VBProjectsEventsGuid, (int)ProjectEventDispId.ItemActivated, _projectActivated);
                _projects = null;
            }
        }

        public static event EventHandler<ProjectEventArgs> ProjectAdded;
        private delegate void ItemAddedDelegate(VB.VBProject vbProject);
        private static ItemAddedDelegate _projectAdded;
        private static void OnProjectAdded(VB.VBProject vbProject)
        {
            if (IsInDesignMode() && vbProject.Protection == VB.vbext_ProjectProtection.vbext_pp_none)
            {
                OnDispatch(ProjectAdded, vbProject, true);
            }
        }

        public static event EventHandler<ProjectEventArgs> ProjectRemoved;
        private delegate void ItemRemovedDelegate(VB.VBProject vbProject);
        private static ItemRemovedDelegate _projectRemoved;
        private static void OnProjectRemoved(VB.VBProject vbProject)
        {
            if (IsInDesignMode() && vbProject.Protection == VB.vbext_ProjectProtection.vbext_pp_none)
            {
                OnDispatch(ProjectRemoved, vbProject);
            }
        }

        public static event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;
        private delegate void ItemRenamedDelegate(VB.VBProject vbProject, string oldName);
        private static ItemRenamedDelegate _projectRenamed;
        private static void OnProjectRenamed(VB.VBProject vbProject, string oldName)
        {
            var project = new VBProject(vbProject);

            if (!IsInDesignMode() || vbProject.Protection == VB.vbext_ProjectProtection.vbext_pp_locked)
            {
                project.Dispose();
                return;
            }

            var projectId = project.ProjectId;

            var handler = ProjectRenamed;
            if (handler == null || projectId == null)
            {
                project.Dispose();
                return;
            }
            handler.Invoke(project, new ProjectRenamedEventArgs(projectId, project, oldName));
        }

        public static event EventHandler<ProjectEventArgs> ProjectActivated;
        private delegate void ItemActivatedDelegate(VB.VBProject vbProject);
        private static ItemActivatedDelegate _projectActivated;
        private static void OnProjectActivated(VB.VBProject vbProject)
        {
            if (IsInDesignMode() && vbProject.Protection == VB.vbext_ProjectProtection.vbext_pp_none)
            {
                OnDispatch(ProjectActivated, vbProject);
            }
        }

        private static void OnDispatch(EventHandler<ProjectEventArgs> dispatched, VB.VBProject vbProject, bool assignId = false)
        {
            var project = new VBProject(vbProject);
            var handler = dispatched;
            if (handler == null || vbProject.Protection == VB.vbext_ProjectProtection.vbext_pp_locked)
            {
                project.Dispose();
                return;
            }

            if (assignId)
            {
                project.AssignProjectId();
            }
            var projectId = project.ProjectId;

            if (projectId == null)
            {
                project.Dispose();
                return;
            }
            handler.Invoke(project, new ProjectEventArgs(projectId, project));
        }

        private static bool IsInDesignMode()
        {
            if (_projects == null)
            {
                return true;
            }
            foreach (var project in _projects.Cast<VB.VBProject>())
            {
                if (project.Mode != VB.vbext_VBAMode.vbext_vm_Design)
                {
                    return false;
                }
            }
            return true;
        }

        #endregion
    }
}