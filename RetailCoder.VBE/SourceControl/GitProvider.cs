using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LibGit2Sharp;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;

namespace Rubberduck.SourceControl
{
    class GitProvider: ISourceControlProvider
    {
        private VBProject project;
        public GitProvider(VBProject project) 
        {
            this.project = project;
            CurrentRepository = new Repository(project.Name, @"C:\Users\Christopher\Documents\SourceControlTest", @"https://github.com/ckuhn203/SourceControlTest.git");
        }

        public Repository CurrentRepository { get; private set; }
        public string CurrentBranch 
        { 
            get 
            {
                using (var repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation))
                {
                    LibGit2Sharp.Branch branch = repo.Branches.Where(b => !b.IsRemote && b.IsCurrentRepositoryHead).First();
                    return branch.Name;
                }
            } 
        }

        public IEnumerable<string> Branches
        {
            get 
            {
                using (var repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation))
                {
                    return repo.Branches.Where(b => !b.IsRemote)
                                        .Select(b => b.Name);
                }

            }
        }

        public static Repository Clone(string remotePathOrUrl, string workingDirectory)
        {
            LibGit2Sharp.Repository.Clone(remotePathOrUrl, workingDirectory);
            //todo: parse name from remote path
            return new Repository(String.Empty, workingDirectory, remotePathOrUrl);
        }

        public static Repository Init(string directory, Microsoft.Vbe.Interop.VBProject project)
        {
            throw new NotImplementedException();
        }

        public void Push()
        {
            using (var repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation))
            {
                repo.Network.Push(repo.Branches[this.CurrentBranch]);
            }
        }

        public void Fetch()
        {
            throw new NotImplementedException();
        }

        public void Pull()
        {
            throw new NotImplementedException();
        }

        public void Commit(string message)
        {
            this.project.ExportSourceFiles(this.CurrentRepository.LocalLocation);

            using (var repo = new LibGit2Sharp.Repository(this.CurrentRepository.LocalLocation))
            {
                RepositoryStatus status = repo.RetrieveStatus();
                List<string> filePaths = status.Modified.Select(mods => mods.FilePath).ToList();
                repo.Stage(filePaths);
                repo.Commit(message);
            }
        }

        public void Merge(string sourceBranch, string destinationBranch)
        {
            using (var repo = new LibGit2Sharp.Repository(this.CurrentRepository.LocalLocation))
            {
                repo.Checkout(repo.Branches[destinationBranch]);
                //todo: find a way to grab this info from git config
                Signature sig = new Signature("ckuhn203", "ckuhn203@gmail.com", System.DateTimeOffset.Now);
                repo.Merge(repo.Branches[sourceBranch], sig);
            }
        }

        public void Checkout(string branch)
        {
            using (var repo = new LibGit2Sharp.Repository(this.CurrentRepository.LocalLocation))
            {
                repo.Checkout(repo.Branches[branch]);
            }

            this.project.RemoveAllComponents();

            var dirInfo = new System.IO.DirectoryInfo(this.CurrentRepository.LocalLocation);
            foreach(var file in dirInfo.EnumerateFiles())
            {
                this.project.VBComponents.ImportSourceFile(file.FullName);
            }
        }

        public void Undo(string filePath)
        {
            // I'll need to parse the component name from the file name
            // Remove the component from the project
            // then reload it from the file system
            throw new NotImplementedException();
        }

        public void Revert()
        {
            throw new NotImplementedException();
        }

        public void AddFile(string filePath)
        {
            throw new NotImplementedException();
        }

        public void RemoveFile(string filePath)
        {
            throw new NotImplementedException();
        }
    }
}
