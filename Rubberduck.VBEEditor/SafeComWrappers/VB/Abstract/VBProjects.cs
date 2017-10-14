using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.Abstract
{
    public abstract class VBProjects
    {
        private static readonly object _locker = new object();
        private static object _projects;
        private static IComIds _comIds;        

        protected VBProjects(object target, VBType vbType)
        {
            lock (_locker)
            {
                if (_projects != null || target == null) { return; }

                _comIds = ComIds.For[vbType];
                AttachEvents(target);
            }
        }

        protected delegate void ProjectAddedDelegate(object vbProject);
        protected static ProjectAddedDelegate _projectAdded;
        public static event EventHandler<ProjectEventArgs> ProjectAdded;
        protected static void DispatchProjectAdded(object vbProject)
        {
            Dispatch(ProjectAdded, vbProject);
        }

        protected delegate void ProjectRemovedDelegate(object vbProject);
        protected static ProjectRemovedDelegate _projectRemoved;
        public static event EventHandler<ProjectEventArgs> ProjectRemoved;
        protected static void DispatchProjectRemoved(object vbProject)
        {
            Dispatch(ProjectRemoved, vbProject);
        }

        protected delegate void ProjectRenamedDelegate(object vbProject, string oldName);
        protected static ProjectRenamedDelegate _projectRenamed;
        public static event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;
        protected static void DispatchProjectRenamed(object vbProject, string oldName)
        {
            Dispatch(ProjectRenamed, vbProject, oldName);
        }

        protected delegate void ProjectActivatedDelegate(object vbProject);
        protected static ProjectActivatedDelegate _projectActivated;
        public static event EventHandler<ProjectEventArgs> ProjectActivated;
        protected static void DispatchProjectActivated(object vbProject)
        {
            Dispatch(ProjectActivated, vbProject);
        }


        private static void Dispatch(EventHandler<ProjectEventArgs> handler, object vbProject)
        {
            var localHandler = handler;
            if (localHandler != null)
            {                
                var project = VBProjectFactory.Create(vbProject);
                if (project.Protection != ProjectProtection.Locked)
                {
                    localHandler.Invoke(project, new ProjectEventArgs(project.ProjectId, project));
                }
            }
        }

        private static void Dispatch(EventHandler<ProjectRenamedEventArgs> handler, object vbProject, string oldName)
        {
            var localHandler = handler;
            if (localHandler != null)
            {
                var project = VBProjectFactory.Create(vbProject);
                if (project.Protection != ProjectProtection.Locked)
                {
                    localHandler.Invoke(project, new ProjectRenamedEventArgs(project.ProjectId, project, oldName));
                }
            }
        }

        private static void AttachEvents(object projects)
        {            
            _projects = projects;
            _projectAdded = DispatchProjectAdded;
            _projectRemoved = DispatchProjectRemoved;
            _projectRenamed = DispatchProjectRenamed;
            _projectActivated = DispatchProjectActivated;
            ComEventsHelper.Combine(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemAdded, _projectAdded);
            ComEventsHelper.Combine(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemRemoved, _projectRemoved);
            ComEventsHelper.Combine(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemRenamed, _projectRenamed);
            ComEventsHelper.Combine(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemActivated, _projectActivated);         
        }

        public static void DetachEvents()
        {
            if (_projects != null)
            {
                ComEventsHelper.Remove(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemAdded, _projectAdded);
                ComEventsHelper.Remove(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemRemoved, _projectRemoved);
                ComEventsHelper.Remove(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemRenamed, _projectRenamed);
                ComEventsHelper.Remove(_projects, _comIds.VBProjectsEventsGuid, _comIds.ProjectEventDispIds.ItemActivated, _projectActivated);
                _projects = null;
            }
        }
    }

    public abstract class VBProjects<T> : VBProjects, ISafeComWrapper<T>, IVBProjects
        where T : class
    {
        private readonly VBProjectsWrapper<T> _comWrapper;
        protected VBProjects(T target, VBType vbType) 
            : base(target, vbType)
        {
            _comWrapper = new VBProjectsWrapper<T>(target, Equals, GetHashCode);
        }


        public abstract IEnumerator<IVBProject> GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public abstract int Count { get; }
        public abstract IVBProject this[object index] { get; }
        public abstract IVBE VBE { get; }
        public abstract IVBE Parent { get; }
        public abstract IVBProject Open(string path);
        public abstract void Remove(IVBProject item);
        public abstract IVBProject Add(ProjectType type);        
        public abstract bool Equals(IVBProjects other);       
        public abstract override int GetHashCode();
        protected bool IsEqualIfNull(ISafeComWrapper<T> other) => _comWrapper.DoIsEqualIfNull(other);
        public T Target => _comWrapper.Target;
        public bool IsWrappingNullReference => _comWrapper.IsWrappingNullReference;

        private class VBProjectsWrapper<TItem> : SafeComWrapper<TItem>
            where TItem : class
        {
            private readonly Func<ISafeComWrapper<TItem>, bool> _equals;
            private readonly Func<int> _getHashCode;
            internal VBProjectsWrapper(TItem target, Func<ISafeComWrapper<TItem>, bool> equals, Func<int> getHashCode)
                : base(target)
            {
                _equals = equals;
                _getHashCode = getHashCode;
            }

            public override bool Equals(ISafeComWrapper<TItem> other)
            {
                return _equals.Invoke(other);
            }

            public override int GetHashCode()
            {
                return _getHashCode.Invoke();
            }

            internal bool DoIsEqualIfNull(ISafeComWrapper<TItem> other) => IsEqualIfNull(other);
        }

        object INullObjectWrapper.Target => _comWrapper.Target;
    }
}
