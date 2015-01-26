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
        public SourceControlProviderBase(VBProject project, Repository repository)
        {
            this.project = project;
            //CurrentRepository = new Repository(project.Name, @"C:\Users\Christopher\Documents\SourceControlTest", @"https://github.com/ckuhn203/SourceControlTest.git");
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

            return new Repository(project.Name, directory, null);
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
            // I'll need to parse the component name from the file name
            // Remove the component from the project
            // then reload it from the file system

            throw new NotImplementedException();
        }

        public virtual void Revert()
        {
            Refresh();
        }

        private void Refresh()
        {
            this.project.RemoveAllComponents();

            var dirInfo = new System.IO.DirectoryInfo(this.CurrentRepository.LocalLocation);
            foreach (var file in dirInfo.EnumerateFiles())
            {
                this.project.VBComponents.ImportSourceFile(file.FullName);
            }
        }
    }
}
