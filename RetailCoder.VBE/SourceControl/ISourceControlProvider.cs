using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl
{
    public interface ISourceControlProvider
    {
        IRepository CurrentRepository { get; }
        IBranch CurrentBranch { get; }
        IEnumerable<IBranch> Branches { get; }

        /// <summary>Clone a remote repository.</summary>
        /// <param name="remotePathOrUrl">Either a Url "https://github.com/retailcoder/Rubberduck.git" or a UNC path. "//server/share/path/to/repo.git"</param>
        /// <param name="workingDirectory">Directory the repository will be cloned to.</param>
        /// <returns>Newly cloned repository.</returns>
        IRepository Clone(string remotePathOrUrl, string workingDirectory);

        /// <summary>
        /// Creates a new repository in/from the given directory.
        /// </summary>
        /// <param name="directory">The directory where the new repository will be created.</param>
        /// <returns>Newly created repository.</returns>
        IRepository Init(string directory, bool bare = false);

        /// <summary>
        /// Creates a new repository and sets the CurrentRepository property from the VBProject passed to the ISourceControlProvider upon creation.
        /// </summary>
        /// <param name="directory"></param>
        /// <returns>Newly created Repository.</returns>
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
        void RemoveFile(string filePath);

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
    }
}
