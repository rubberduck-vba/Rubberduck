using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.NonDisposalDecorators
{
    public class VBProjectsNonDisposalDecorator<T> : NonDisposalDecoratorBase<T>, IVBProjects
        where T : IVBProjects
    {
        public VBProjectsNonDisposalDecorator(T projects)
            : base(projects)
        { }

        public void AttachEvents()
        {
            WrappedItem.AttachEvents();
        }

        public void DetachEvents()
        {
            WrappedItem.DetachEvents();
        }

        public IEnumerator<IVBProject> GetEnumerator()
        {
            return WrappedItem.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable) WrappedItem).GetEnumerator();
        }

        public int Count => WrappedItem.Count;

        public IVBProject this[object index] => WrappedItem[index];

        public bool Equals(IVBProjects other)
        {
            return WrappedItem.Equals(other);
        }

        public event EventHandler<ProjectEventArgs> ProjectActivated
        {
            add => WrappedItem.ProjectActivated += value;
            remove => WrappedItem.ProjectActivated -= value;
        }

        public event EventHandler<ProjectEventArgs> ProjectAdded
        {
            add => WrappedItem.ProjectAdded += value;
            remove => WrappedItem.ProjectAdded -= value;
        }

        public event EventHandler<ProjectEventArgs> ProjectRemoved
        {
            add => WrappedItem.ProjectRemoved += value;
            remove => WrappedItem.ProjectRemoved -= value;
        }

        public event EventHandler<ProjectRenamedEventArgs> ProjectRenamed
        {
            add => WrappedItem.ProjectRenamed += value;
            remove => WrappedItem.ProjectRenamed -= value;
        }

        public IVBE VBE => WrappedItem.VBE;

        public IVBE Parent => WrappedItem.Parent;

        public IVBProject Add(ProjectType type)
        {
            return WrappedItem.Add(type);
        }

        public IVBProject Open(string path)
        {
            return WrappedItem.Open(path);
        }

        public void Remove(IVBProject project)
        {
            WrappedItem.Remove(project);
        }

        public IVBProject StartProject
        {
            get => WrappedItem.StartProject;
            set => WrappedItem.StartProject = value;
        }
    }
}