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
    class GitProvider : SourceControlProviderBase
    {
        public GitProvider(VBProject project, Repository repository)
            :base(project, repository) {}

        public override string CurrentBranch 
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

        public override IEnumerable<string> Branches
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

        public override Repository Clone(string remotePathOrUrl, string workingDirectory)
        {
            LibGit2Sharp.Repository.Clone(remotePathOrUrl, workingDirectory);
            //todo: parse name from remote path
            return new Repository(String.Empty, workingDirectory, remotePathOrUrl);
        }

        public override Repository Init(string directory, Microsoft.Vbe.Interop.VBProject project)
        {
            throw new NotImplementedException();
        }

        public override void Push()
        {
            using (var repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation))
            {
                repo.Network.Push(repo.Branches[this.CurrentBranch]);
            }
        }

        public override void Fetch()
        {
            throw new NotImplementedException();
        }

        public override void Pull()
        {
            throw new NotImplementedException();
        }

        public override void Commit(string message)
        {
            base.Commit(message);

            using (var repo = new LibGit2Sharp.Repository(this.CurrentRepository.LocalLocation))
            {
                RepositoryStatus status = repo.RetrieveStatus();
                List<string> filePaths = status.Modified.Select(mods => mods.FilePath).ToList();
                repo.Stage(filePaths);
                repo.Commit(message);
            }
        }

        public override void Merge(string sourceBranch, string destinationBranch)
        {
            using (var repo = new LibGit2Sharp.Repository(this.CurrentRepository.LocalLocation))
            {
                repo.Checkout(repo.Branches[destinationBranch]);
                //todo: find a way to grab this info from git config
                Signature sig = new Signature("ckuhn203", "ckuhn203@gmail.com", System.DateTimeOffset.Now);
                repo.Merge(repo.Branches[sourceBranch], sig);
            }
        }

        public override void Checkout(string branch)
        {
            using (var repo = new LibGit2Sharp.Repository(this.CurrentRepository.LocalLocation))
            {
                repo.Checkout(repo.Branches[branch]);
            }

            base.Checkout(branch);
        }

        //public void Undo(string filePath)
        //{

        //    throw new NotImplementedException();
        //}

        public override void Revert()
        {
            throw new NotImplementedException();
        }

        public override void AddFile(string filePath)
        {
            throw new NotImplementedException();
        }

        public override void RemoveFile(string filePath)
        {
            throw new NotImplementedException();
        }
    }
}
