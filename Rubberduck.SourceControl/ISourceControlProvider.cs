using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl
{
    public interface ISourceControlProvider
    {
        IRepository CurrentRepository { get; }
        IBranch CurrentBranch { get; }
        IEnumerable<IBranch> Branches { get; }
        IList<ICommit> UnsyncedLocalCommits { get; }
        IList<ICommit> UnsyncedRemoteCommits { get; }
        bool NotifyExternalFileChanges { get; set; }
        bool HandleVbeSinkEvents { get; set; }

        event EventHandler<EventArgs> BranchChanged;

        /// <summary>Clone a remote repository.</summary>
        /// <param name="remotePathOrUrl">Either a Url "https://github.com/retailcoder/Rubberduck.git" or a UNC path. "//server/share/path/to/repo.git"</param>
        /// <param name="workingDirectory">Directory the repository will be cloned to.</param>
        /// <param name="credentials">Credentials required if repository is private.</param>
        /// <returns>Newly cloned repository.</returns>
        IRepository Clone(string remotePathOrUrl, string workingDirectory, SecureCredentials credentials = null);

        /// <summary>
        /// Creates a new repository in/from the given directory.
        /// </summary>
        /// <param name="directory">The directory where the new repository will be created.</param>
        /// <param name="bare">Specifies whether or not the new repository will be intialized as a bare repo.</param>
        /// <returns>Newly created repository.</returns>
        IRepository Init(string directory, bool bare = false);

        /// <summary>
        /// Publishes to a remote repo.
        /// </summary>
        /// <param name="path">The remote path to the repo</param>
        /// <param name="trackingBranchName">The branch name to publish to</param>
        void AddOrigin(string path, string trackingBranchName);

        /// <summary>
        /// Creates a new repository and sets the CurrentRepository property from the VBProject passed to the ISourceControlProvider upon creation.
        /// </summary>
        /// <param name="directory"></param>
        /// <returns>Newly created Repository.</returns>
        // ReSharper disable once InconsistentNaming : Changing this now will break the COM interface.
        IRepository InitVBAProject(string directory);

        /// <summary>
        /// Pushes commits in the CurrentBranch of the Local repo to the Remote.
        /// </summary>
        void Push();

        /// <summary>
        /// Fetches the specified remote for tracking.
        /// If argument is not supplied, returns a default remote defined by implementation.
        /// </summary>
        /// <param name="remoteName">Name of the remote to be fetched.</param>
        void Fetch([Optional] string remoteName);

        /// <summary>
        /// Fetches the currently tracking remote and merges it into the CurrentBranch.
        /// </summary>
        void Pull();

        /// <summary>
        /// Stages all modified files and commits to CurrentBranch.
        /// </summary>
        /// <param name="message">Commit message.</param>
        void Commit(string message);

        /// <summary>
        /// Merges the source branch into the desitnation.
        /// </summary>
        /// <param name="sourceBranch">Name of the source branch.</param>
        /// <param name="destinationBranch">Name of the target branch.</param>
        void Merge(string sourceBranch, string destinationBranch);

        /// <summary>
        /// Checks out the target branch.
        /// </summary>
        /// <param name="branch">Name of the branch to be checked out.</param>
        void Checkout(string branch);

        /// <summary>
        /// Creates and checks out a new branch.
        /// </summary>
        /// <param name="branch">Name of the branch to be created.</param>
        void CreateBranch(string branch);

        /// <summary>
        /// Creates and checks out a new branch.
        /// </summary>
        /// <param name="sourceBranch">Name of the branch to base the new branch on.</param>
        /// <param name="branch">Name of the branch to be created.</param>
        void CreateBranch(string sourceBranch, string branch);

        /// <summary>
        /// Deletes the specified branch from the local repository.
        /// </summary>
        /// <param name="branch">Name of the branch to be deleted.</param>
        void DeleteBranch(string branch);

        /// <summary>
        /// Undoes uncommitted changes to a particular file.
        /// </summary>
        /// <param name="filePath"></param>
        void Undo(string filePath);

        /// <summary>
        /// Reverts entire branch to the last commit.
        /// </summary>
        void Revert();

        /// <summary>
        /// Adds untracked file to repository.
        /// </summary>
        /// <param name="filePath"></param>
        void AddFile(string filePath);

        /// <summary>
        /// Removes file from tracking.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="removeFromWorkingDirectory"></param>
        void RemoveFile(string filePath, bool removeFromWorkingDirectory);

        /// <summary>
        /// Returns a collection of file status entries.
        /// Semantically the same as calling $git status.
        /// </summary>
        IEnumerable<IFileStatusEntry> Status();

        /// <summary>
        /// Stages a file to be committed.
        /// </summary>
        /// <param name="filePath"></param>
        void Stage(string filePath);

        /// <summary>
        /// Stages a list of files to be committed.
        /// </summary>
        /// <param name="filePaths"></param>
        void Stage(IEnumerable<string> filePaths);

        /// <summary>
        /// Publish a local branch.
        /// </summary>
        /// <param name="branch">The name of the branch to publish</param>
        void Publish(string branch);

        /// <summary>
        /// Unpublish a remote branch.
        /// </summary>
        /// <param name="branch">The name of the branch to unpublish</param>
        void Unpublish(string branch);

        /// <summary>
        /// Returns whether user has credentials to log into credentials.
        /// </summary>
        /// <returns>Returns true if repo can log into GitHub.</returns>
        bool HasCredentials();

        /// <summary>
        /// Gets the last known status without refreshing
        /// </summary>
        /// <returns>Collection of statuses.</returns>
        IEnumerable<IFileStatusEntry> LastKnownStatus();

        /// <summary>
        /// Reloads the component into the VBE
        /// </summary>
        /// <param name="fileName"></param>
        void ReloadComponent(string fileName);

        /// <summary>
        /// Returns whether the repo has a remote named "origin"
        /// </summary>
        /// <returns></returns>
        bool RepoHasRemoteOrigin();
    }
}
