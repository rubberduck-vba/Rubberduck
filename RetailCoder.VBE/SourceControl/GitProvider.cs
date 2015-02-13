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
    //todo: I need a way to get to the branch status...
    class GitProvider : SourceControlProviderBase
    {
        private LibGit2Sharp.Repository repo;
        //todo: put this someplace safe and preferably encrypted
        private Credentials creds;
        private LibGit2Sharp.Handlers.CredentialsHandler credHandler;

        public GitProvider(VBProject project, Repository repository, string userName, string passWord)
            : base(project, repository)
        {
            repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation);

            this.creds = new UsernamePasswordCredentials()
            {
                Username = userName,
                Password = passWord
            };

            this.credHandler = (url, user, cred) => creds;
        }

        ~GitProvider()
        {
            repo.Dispose();
        }

        public override string CurrentBranch
        {
            get
            {
                return repo.Branches.Where(b => !b.IsRemote && b.IsCurrentRepositoryHead)
                                    .First().Name;
            }
        }

        public override IEnumerable<string> Branches
        {
            get
            {
                return repo.Branches.Where(b => !b.IsRemote)
                                    .Select(b => b.Name);
            }
        }

        //note: should we really have a clone method? I don't think we'll use it, but could be useful in an API.
        public override Repository Clone(string remotePathOrUrl, string workingDirectory)
        {
            //todo: parse name from remote path
            var name = String.Empty;
            return new Repository(name, workingDirectory, remotePathOrUrl);
        }

        public override Repository Init(string directory)
        {
            //todo: implement
            var repository = base.Init(directory);

            throw new NotImplementedException();

            return repository;
        }

        public override void Push()
        {
            try
            {
                var options = new PushOptions()
                {
                    CredentialsProvider = credHandler
                };

                var branch = repo.Branches[this.CurrentBranch];
                repo.Network.Push(branch, options);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Push Failed.", ex);
            }
        }

        /// <summary>
        /// Fetches the specified remote for tracking.
        /// If not argument is supplied, fetches the "origin" remote.
        /// </summary>
        public override void Fetch()
        {
            Fetch("origin");
        }

        public void Fetch(string remoteName)
        {
            try
            {
                var remote = repo.Network.Remotes[remoteName];
                repo.Network.Fetch(remote);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Fetch failed.", ex);
            }
        }

        public override void Pull()
        {
            try
            {

                var options = new PullOptions()
                {
                    MergeOptions = new MergeOptions()
                    {
                        FastForwardStrategy = FastForwardStrategy.Default
                    }
                };

                var signature = GetSignature();
                repo.Network.Pull(signature, options);

                base.Pull();
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Pull Failed.", ex);
            }
        }

        public override void Commit(string message)
        {
            try
            {
                base.Commit(message);

                RepositoryStatus status = repo.RetrieveStatus();
                List<string> filePaths = status.Modified.Select(mods => mods.FilePath).ToList();
                repo.Stage(filePaths);
                repo.Commit(message);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Commit Failed.", ex);
            }
        }

        public override void Merge(string sourceBranch, string destinationBranch)
        {
            repo.Checkout(repo.Branches[destinationBranch]);

            var oldHeadCommit = repo.Head.Tip;

            Signature signature = GetSignature();
            var result = repo.Merge(repo.Branches[sourceBranch], signature);

            switch (result.Status)
            {
                case MergeStatus.Conflicts:
                    repo.Reset(ResetMode.Hard, oldHeadCommit); //abort merge
                    //todo: report conflict? Perhaps merge should follow Lib2Git's lead and return a value?
                    break;
                case MergeStatus.NonFastForward:
                    //https://help.github.com/articles/dealing-with-non-fast-forward-errors/
                    Pull();
                    Merge(sourceBranch, destinationBranch); //a little leary about this. Could stack overflow if I'm wrong.
                    break;
                default:
                    break; //success
            }

            base.Merge(sourceBranch, destinationBranch);
        }

        public override void Checkout(string branch)
        {
            repo.Checkout(repo.Branches[branch]);

            base.Checkout(branch);
        }

        public override void CreateBranch(string branch)
        {
            try
            {
                repo.CreateBranch(branch);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Branch creation failed.", ex);
            }
        }

        public override void Revert()
        {
            //todo: investigate revert results class
            repo.Revert(repo.Head.Tip, GetSignature());

            base.Revert();
        }

        public override void AddFile(string filePath)
        {
            // https://github.com/libgit2/libgit2sharp/wiki/Git-add
            repo.Stage(filePath);
        }

        public override void RemoveFile(string filePath)
        {
            //todo: implement
            throw new NotImplementedException();
        }

        private Signature GetSignature()
        {
            return this.repo.Config.BuildSignature(DateTimeOffset.Now);
            //todo: clean up dead code
            // return new Signature("ckuhn203", "ckuhn203@gmail.com", DateTimeOffset.Now);
        }
    }
}
