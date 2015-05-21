using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using LibGit2Sharp;
using LibGit2Sharp.Handlers;
using Microsoft.Vbe.Interop;
using System.Security;

namespace Rubberduck.SourceControl
{
    public class GitProvider : SourceControlProviderBase
    {
        private readonly LibGit2Sharp.Repository _repo;
        private readonly LibGit2Sharp.Credentials _credentials;
        private readonly CredentialsHandler _credentialsHandler;

        public GitProvider(VBProject project) 
            : base(project) { }

        public GitProvider(VBProject project, IRepository repository)
            : base(project, repository) 
        {
            try
            {
                _repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation);
            }
            catch (RepositoryNotFoundException ex)
            {
                throw new SourceControlException("Repository not found.", ex);
            }
        }
        
        public GitProvider(VBProject project, IRepository repository, string userName, string passWord)
            : this(project, repository)
        {
            _credentials = new UsernamePasswordCredentials()
            {
                Username = userName,
                Password = passWord
            };

            _credentialsHandler = (url, user, cred) => _credentials;
        }

        public GitProvider(VBProject project, IRepository repository, ICredentials<string> credentials)
            :this(project, repository, credentials.Username, credentials.Password)
        { }

        public GitProvider(VBProject project, IRepository repository, ICredentials<SecureString> credentials)
            : this(project, repository)
        {
            _credentials = new SecureUsernamePasswordCredentials()
            {
                Username = credentials.Username,
                Password = credentials.Password
            };
        }

        ~GitProvider()
        {
            if (_repo != null)
            {
                _repo.Dispose();
            }
        }

        public override IBranch CurrentBranch
        {
            get
            {
                return this.Branches.First(b => !b.IsRemote && b.IsCurrentHead);
            }
        }

        public override IEnumerable<IBranch> Branches
        {
            get
            {
                //note: consider doing this once and refreshing if necessary
                return _repo.Branches.Select(b => new Branch(b));
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
                var workingDir = (bare) ? string.Empty : directory;

                LibGit2Sharp.Repository.Init(directory, bare);

                return new Repository(this.project.Name, workingDir, directory);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Unable to initialize repository.", ex);
            }
        }

        /// <summary>
        /// Exports files from VBProject to the file system, initalizes the repository, and creates an inital commit of those files to the repo.
        /// </summary>
        /// <param name="directory">Local file path of the directory where the new repository will be created.</param>
        /// <returns>Newly initialized repository.</returns>
        public override IRepository InitVBAProject(string directory)
        {
            var repository = base.InitVBAProject(directory);
            Init(repository.LocalLocation);

            //add a master branch to newly created repo
            using (var repo = new LibGit2Sharp.Repository(repository.LocalLocation))
            {
                var status = repo.RetrieveStatus(new StatusOptions());
                foreach (var stat in status.Untracked)
                {
                    repo.Stage(stat.FilePath);
                }

                repo.Commit("Intial Commit");
            }

            return repository;
        }

        public override void Push()
        {
            try
            {
                //Only use credentials if we've been given credentials to use in the constructor.
                PushOptions options = null;
                if (_credentials != null)
                {
                    options = new PushOptions()
                    {
                        CredentialsProvider = _credentialsHandler
                    };
                }

                var branch = _repo.Branches[this.CurrentBranch.Name];
                _repo.Network.Push(branch, options);
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
                var remote = _repo.Network.Remotes[remoteName];
                _repo.Network.Fetch(remote);
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
                _repo.Network.Pull(signature, options);

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
                _repo.Commit(message);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Commit Failed.", ex);
            }
        }

        public override void Stage(string filePath)
        {
            try
            {
                _repo.Stage(filePath);
            }
            catch (LibGit2SharpException ex)
            {
                throw  new SourceControlException("Failed to stage file.", ex);
            }
        }

        public override void Stage(IEnumerable<string> filePaths)
        {
            try
            {
                _repo.Stage(filePaths);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Failed to stage file.", ex);
            }
        }

        public override void Merge(string sourceBranch, string destinationBranch)
        {
            _repo.Checkout(_repo.Branches[destinationBranch]);

            var oldHeadCommit = _repo.Head.Tip;
            var signature = GetSignature();
            var result = _repo.Merge(_repo.Branches[sourceBranch], signature);

            switch (result.Status)
            {
                case MergeStatus.Conflicts:
                    _repo.Reset(ResetMode.Hard, oldHeadCommit);
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
                _repo.Checkout(_repo.Branches[branch]);
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
                _repo.CreateBranch(branch);
                _repo.Checkout(branch);
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
                var results = _repo.Revert(_repo.Head.Tip, GetSignature());

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
                _repo.Stage(filePath);
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
                _repo.Remove(filePath, false);
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
                return _repo.RetrieveStatus().Select(item => new FileStatusEntry(item));
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
                _repo.CheckoutPaths(this.CurrentBranch.Name, new List<string> {filePath});
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
                if (_repo.Branches.Any(b => b.Name == branch && !b.IsRemote))
                {
                    _repo.Branches.Remove(branch);
                }
            }
            catch(LibGit2SharpException ex)
            {
                throw new SourceControlException("Branch deletion failed.", ex);
            }
        }

        private Signature GetSignature()
        {
            return _repo.Config.BuildSignature(DateTimeOffset.Now);
        }
    }
}
