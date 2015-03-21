using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;

namespace Rubberduck.SourceControl
{
    public abstract class SourceControlProviderBase : ISourceControlProvider
    {
        protected VBProject project;

        public SourceControlProviderBase(VBProject project)
        {
            this.project = project;
        }

        public SourceControlProviderBase(VBProject project, IRepository repository)
            :this(project)
        {
            this.CurrentRepository = repository;
        }

        public IRepository CurrentRepository { get; private set; }
        public abstract string CurrentBranch { get; }
        public abstract IEnumerable<string> Branches { get; }
        public abstract IRepository Clone(string remotePathOrUrl, string workingDirectory);
        public abstract void Push();
        public abstract void Fetch(string remoteName);
        public abstract void AddFile(string filePath);
        public abstract void RemoveFile(string filePath);
        public abstract void CreateBranch(string branch);
        public abstract void DeleteBranch(string branch);
        public abstract IRepository Init(string directory, bool bare = false);

        public virtual IRepository InitVBAProject(string directory)
        {
            var projectName = GetProjectNameFromDirectory(directory);
            if (projectName != string.Empty && projectName != this.project.Name)
            {
                directory = Path.Combine(directory, project.Name);
            }

            this.project.ExportSourceFiles(directory);
            this.CurrentRepository = new Repository(project.Name, directory, directory);
            return this.CurrentRepository;
        }

        public virtual void Pull()
        {
            Refresh();
        }

        public virtual void Commit(string message)
        {
            this.project.ExportSourceFiles(this.CurrentRepository.LocalLocation);
        }

        public virtual void Merge(string sourceBranch, string destinationBranch)
        {
            Refresh();
        }

        public virtual void Checkout(string branch)
        {
            Refresh();
        }

        public virtual void Undo(string filePath)
        {
            //this might need to cherry pick from the tip instead.

           var componentName = Path.GetFileNameWithoutExtension(filePath);

           //GetFileNameWithoutExtension returns empty string if it's not a file
           //https://msdn.microsoft.com/en-us/library/system.io.path.getfilenamewithoutextension%28v=vs.110%29.aspx
            if (componentName != String.Empty)
            {
                var component = this.project.VBComponents.Item(componentName);
                this.project.VBComponents.RemoveSafely(component);
                this.project.VBComponents.ImportSourceFile(filePath);
            }
        }

        public virtual void Revert()
        {
            Refresh();
        }

        public virtual IEnumerable<IFileStatusEntry> Status()
        {
            this.project.ExportSourceFiles(this.CurrentRepository.LocalLocation);
            return null;
        }

        protected string GetProjectNameFromDirectory(string directory)
        {
            var separators = new char[] { '/', '\\', '.' };
            return directory.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                .Where(c => c != "git")
                .LastOrDefault();
        }

        private void Refresh()
        {
            //Because refreshing removes all components, we need to store the current selection,
            // so we can correctly reset it once the files are imported from the repository.
            var selection = project.VBE.ActiveCodePane.GetSelection();

            project.RemoveAllComponents();
            project.ImportSourceFiles(CurrentRepository.LocalLocation);

            project.VBE.SetSelection(selection);
        }
    }
}
