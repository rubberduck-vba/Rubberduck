using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using LibGit2Sharp;
using LibGit2Sharp.Handlers;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.SourceControl
{
    public class GitProvider : SourceControlProviderBase
    {
        private readonly LibGit2Sharp.Repository _repo;
        private readonly LibGit2Sharp.Credentials _credentials;
        private readonly CredentialsHandler _credentialsHandler;
        private List<ICommit> _unsyncedLocalCommits;
        private List<ICommit> _unsyncedRemoteCommits;

        public GitProvider(VBProject project)
            : base(project)
        {
            _unsyncedLocalCommits = new List<ICommit>();
            _unsyncedRemoteCommits = new List<ICommit>();
        }

        public GitProvider(VBProject project, IRepository repository, ICodePaneWrapperFactory wrapperFactory)
            : base(project, repository, wrapperFactory) 
        {
            _unsyncedLocalCommits = new List<ICommit>();
            _unsyncedRemoteCommits = new List<ICommit>();

            try
            {
                _repo = new LibGit2Sharp.Repository(CurrentRepository.LocalLocation);
            }
            catch (RepositoryNotFoundException ex)
            {
                throw new SourceControlException("Repository not found.", ex);
            }
        }

        public GitProvider(VBProject project, IRepository repository, string userName, string passWord, ICodePaneWrapperFactory wrapperFactory)
            : this(project, repository, wrapperFactory)
        {
            _credentials = new UsernamePasswordCredentials()
            {
                Username = userName,
                Password = passWord
            };

            _credentialsHandler = (url, user, cred) => _credentials;
        }

        public GitProvider(VBProject project, IRepository repository, ICredentials<string> credentials, ICodePaneWrapperFactory wrapperFactory)
            :this(project, repository, credentials.Username, credentials.Password, wrapperFactory)
        { }

        public GitProvider(VBProject project, IRepository repository, ICredentials<SecureString> credentials, ICodePaneWrapperFactory wrapperFactory)
            : this(project, repository, wrapperFactory)
        {
            _credentials = new SecureUsernamePasswordCredentials()
            {
                Username = credentials.Username,
                Password = credentials.Password
            };

            _credentialsHandler = (url, user, cred) => _credentials;
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
                return Branches.First(b => !b.IsRemote && b.IsCurrentHead);
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

        public override IList<ICommit> UnsyncedLocalCommits
        {
            get { return _unsyncedLocalCommits; }
        }

        public override IList<ICommit> UnsyncedRemoteCommits
        {
            get { return _unsyncedRemoteCommits; }
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

                return new Repository(Project.Name, workingDir, directory);
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

                try
                {
                    //The default behavior of LibGit2Sharp.Repo.Commit is to throw an exception if no signature is found,
                    // but BuildSignature() does not throw if a signature is not found, it returns "unknown" instead.
                    // so we pass a signature that won't throw along to the commit.
                    repo.Commit("Intial Commit", GetSignature(repo));
                }
                catch(LibGit2SharpException ex)
                {
                    throw new SourceControlException("Unable to perform intial commit.", ex);
                }
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

                var branch = _repo.Branches[CurrentBranch.Name];
                _repo.Network.Push(branch, options);

                RequeryUnsyncedCommits();
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

                if (remote != null)
                {
                    _repo.Network.Fetch(remote);
                }

                RequeryUnsyncedCommits();
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

                RequeryUnsyncedCommits();
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
                //The default behavior of LibGit2Sharp.Repo.Commit is to throw an exception if no signature is found,
                // but BuildSignature() does not throw if a signature is not found, it returns "unknown" instead.
                // so we pass a signature that won't throw along to the commit.
                _repo.Commit(message, GetSignature());
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
                    //abort the merge by resetting to the state prior to the merge
                    _repo.Reset(ResetMode.Hard, oldHeadCommit);
                    break;
                case MergeStatus.NonFastForward:
                    //https://help.github.com/articles/dealing-with-non-fast-forward-errors/
                    Pull();
                    Merge(sourceBranch, destinationBranch); //a little leary about this. Could stack overflow if I'm wrong.
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

                RequeryUnsyncedCommits();
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

                RequeryUnsyncedCommits();
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Branch creation failed.", ex);
            }
        }

        public override void CreateBranch(string sourceBranch, string branch)
        {
            try
            {
                _repo.CreateBranch(branch, _repo.Branches[sourceBranch].Commits.Last());
                _repo.Checkout(branch);

                RequeryUnsyncedCommits();
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
                _repo.CheckoutPaths(CurrentBranch.Name, new List<string> {filePath});
                base.Undo(filePath);
            }
            catch (LibGit2SharpException ex)
            {
                throw new SourceControlException("Undo failed.", ex);
            }
        }

        public override void DeleteBranch(string branchName)
        {
            try
            {
                var branch = _repo.Branches.FirstOrDefault(b => b.FriendlyName == branchName);
                if (branch != null)
                {
                    if (branch.IsRemote)
                    {
                        PushOptions options = null;
                        if (_credentials != null)
                        {
                            options = new PushOptions
                            {
                                CredentialsProvider = _credentialsHandler
                            };
                        }

                        _repo.Network.Push(branch.Remote, ":refs/heads/" + branchName.Split('/').Last(),
                            options);
                    }
                    else
                    {
                        _repo.Branches.Remove(branch);
                    }
                }
            }
            catch(LibGit2SharpException ex)
            {
                throw new SourceControlException("Branch deletion failed.", ex);
            }
        }

        private Signature GetSignature(LibGit2Sharp.IRepository repo)
        {
            return repo.Config.BuildSignature(DateTimeOffset.Now);
        }

        private Signature GetSignature()
        {
            return _repo.Config.BuildSignature(DateTimeOffset.Now);
        }

        private void RequeryUnsyncedCommits()
        {
            var currentBranch = _repo.Branches[CurrentBranch.Name];
            var local = currentBranch.Commits;

            if (currentBranch.TrackedBranch == null)
            {
                _unsyncedLocalCommits = local.Select(c => new Commit(c) as ICommit)
                                            .ToList();

                _unsyncedRemoteCommits = new List<ICommit>();
            }
            else
            {
                var remote = currentBranch.TrackedBranch.Commits;

                _unsyncedLocalCommits = local.Where(c => !remote.Contains(c))
                                            .Select(c => new Commit(c) as ICommit)
                                            .ToList();

                _unsyncedRemoteCommits = remote.Where(c => !local.Contains(c))
                                               .Select(c => new Commit(c) as ICommit)
                                               .ToList();
            }
        }
    }
}
