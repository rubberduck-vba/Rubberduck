using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;

namespace Rubberduck.SourceControl
{
    public abstract class SourceControlProviderBase : ISourceControlProvider
    {
        private VBProject project;

        public SourceControlProviderBase(VBProject project)
        {
            this.project = project;
        }

        public SourceControlProviderBase(VBProject project, Repository repository)
            :this(project)
        {
            this.CurrentRepository = repository;
        }

        public Repository CurrentRepository { get; private set; }
        public abstract string CurrentBranch { get; }
        public abstract IEnumerable<string> Branches { get; }
        public abstract Repository Clone(string remotePathOrUrl, string workingDirectory);
        public abstract void Push();
        public abstract void Fetch();
        public abstract void AddFile(string filePath);
        public abstract void RemoveFile(string filePath);


        public virtual Repository Init(string directory)
        {
            this.project.ExportSourceFiles(directory);
            this.CurrentRepository = new Repository(project.Name, directory, String.Empty);
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
            //GetFileNameWithoutExtension returns empty string if it's not a file
            //https://msdn.microsoft.com/en-us/library/system.io.path.getfilenamewithoutextension%28v=vs.110%29.aspx
            var componentName = System.IO.Path.GetFileNameWithoutExtension(filePath);

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

        private void Refresh()
        {
            this.project.RemoveAllComponents();

            var dirInfo = new System.IO.DirectoryInfo(this.CurrentRepository.LocalLocation);

            var files = dirInfo.EnumerateFiles()
                                .Where(f => f.Extension == VBComponentExtensions.StandardExtension||
                                            f.Extension == VBComponentExtensions.ClassExtesnion ||
                                            f.Extension == VBComponentExtensions.DocClassExtension ||
                                            f.Extension == VBComponentExtensions.FormExtension 
                                            );
            foreach (var file in files)
            {
                this.project.VBComponents.ImportSourceFile(file.FullName);
            }
        }
    }
}
