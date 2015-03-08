using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using System.Runtime.InteropServices;
using System.ComponentModel;
using LibGit2Sharp;
using IRepository = Rubberduck.SourceControl.IRepository;

namespace Rubberduck.SourceControl
{
    public class GitProvider : SourceControlProviderBase
    {
        private LibGit2Sharp.Repository repo;
        private Credentials credentials;
        private LibGit2Sharp.Handlers.CredentialsHandler credHandler;

        public GitProvider(VBProject project) 
            : base(project) { }

        public GitProvider(VBProject project, IRepository repository)
            : base(project, repository) 
        {
            repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation);
        }

        public GitProvider(VBProject project, IRepository repository, string userName, string passWord)
            : this(project, repository)
        {
            this.credentials = new UsernamePasswordCredentials()
            {
                Username = userName,
                Password = passWord
            };

            this.credHandler = (url, user, cred) => credentials;
        }

        ~GitProvider()
        {
            if (repo != null)
            {
                repo.Dispose();
            }
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

        public override IRepository Clone(string remotePathOrUrl, string workingDirectory)
        {
            try
            {
                var name = GetProjectNameFromDirectory(remotePathOrUrl);
                LibGit2Sharp.Repository.Clone(remotePathOrUrl, workingDirectory);
                return new Repository(name, workingDirectory, remotePathOrUrl);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Failed to clone remote repository.", ex);
            }
        }

        public override IRepository Init(string directory, bool bare = false)
        {
            try
            {
                string workingDir;

                if (bare)
                {
                    workingDir = string.Empty;
                }
                else
                {
                    workingDir = directory;
                }

                LibGit2Sharp.Repository.Init(directory, bare);

                var projectName = GetProjectNameFromDirectory(directory);
                return new Repository(this.project.Name, workingDir, directory);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Unable to initialize repository.", ex);
            }
        }

        public override IRepository InitVBAProject(string directory)
        {
            var repository = base.InitVBAProject(directory);
            Init(repository.LocalLocation);
            return repository;
        }

        public override void Push()
        {
            try
            {
                //Only use credentials if we've been given credentials to use in the constructor.
                PushOptions options = null;
                if (this.credentials != null)
                {
                    options = new PushOptions()
                    {
                        CredentialsProvider = credHandler
                    };
                }

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
        public override void Fetch([Optional] string remoteName)
        {
            if (remoteName == null)
            {
                remoteName = "origin";
            }

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
            var signature = GetSignature();
            var result = repo.Merge(repo.Branches[sourceBranch], signature);

            switch (result.Status)
            {
                case MergeStatus.Conflicts:
                    repo.Reset(ResetMode.Hard, oldHeadCommit);
                    break;
                case MergeStatus.NonFastForward:
                    //https://help.github.com/articles/dealing-with-non-fast-forward-errors/
                    Pull();
                    Merge(sourceBranch, destinationBranch); //a little leary about this. Could stack overflow if I'm wrong.
                    break;
                default:
                    break;
            }
            base.Merge(sourceBranch, destinationBranch);
        }

        public override void Checkout(string branch)
        {
            try
            {
                repo.Checkout(repo.Branches[branch]);
                base.Checkout(branch);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Checkout failed.", ex);
            }
        }

        public override void CreateBranch(string branch)
        {
            try
            {
                repo.CreateBranch(branch);
                repo.Checkout(branch);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Branch creation failed.", ex);
            }
        }

        public override void Revert()
        {
            try
            {
                var results = repo.Revert(repo.Head.Tip, GetSignature());

                if (results.Status == RevertStatus.Conflicts)
                {
                    throw new SourceControlException("Revert resulted in conflicts. Revert failed.");
                }

                base.Revert();
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Revert failed.", ex);
            }
        }

        public override void AddFile(string filePath)
        {
            try
            {
                // https://github.com/libgit2/libgit2sharp/wiki/Git-add
                repo.Stage(filePath);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException(string.Format("Failed to stage file {0}", filePath), ex);
            }
        }

        /// <summary>
        /// Removes file from staging area, but leaves the file in the working directory.
        /// </summary>
        /// <param name="filePath"></param>
        public override void RemoveFile(string filePath)
        {
            try
            {
                repo.Remove(filePath, false);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException(string.Format("Failed to remove file {0} from staging area.", filePath), ex);
            }
        }

        public override IEnumerable<IFileStatusEntry> Status()
        {
            try
            {
                base.Status();
                return repo.RetrieveStatus().Select(item => new FileStatusEntry(item));
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Failed to retrieve repository status.", ex);
            }
        }

        public override void Undo(string filePath)
        {
            try
            {
                var paths = new List<string>();
                paths.Add(filePath);

                repo.CheckoutPaths(this.CurrentBranch, paths);
                base.Undo(filePath);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Undo failed.", ex);
            }
        }

        public override void DeleteBranch(string branch)
        {
            try
            {
                if (repo.Branches.Any(b => b.Name == branch && !b.IsRemote))
                {
                    repo.Branches.Remove(branch);
                }
            }
            catch(LibGit2SharpException ex)
            {
                throw new SourceControlException("Branch deletion failed.", ex);
            }
        }

        private Signature GetSignature()
        {
            return this.repo.Config.BuildSignature(DateTimeOffset.Now);
        }
    }
}
