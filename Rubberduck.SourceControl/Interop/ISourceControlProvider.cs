using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl.Interop
{
    [ComVisible(true)]
    [Guid("A44AF849-3C48-4303-A855-4B156958F3C4")]
    public interface ISourceControlProvider
    {
        [DispId(0)]
        IRepository CurrentRepository { get; }

        [DispId(1)]
        string CurrentBranch { get; }

        [DispId(2)]
        IEnumerable Branches { get; }

        [DispId(3)]
        [Description("Clones a remote repository to the local file system.")]
        IRepository Clone(string remotePathOrUrl, string workingDirectory, Credentials credentials = null);

        [DispId(4)]
        [Description("Creates a new repository in/from the specified directory.")]
        IRepository Init(string directory, bool bare = false);

        [DispId(5)]
        [Description("Creates a new repository and sets the CurrentRepository property from the VBProject passed to the ISourceControlProvider upon creation.")]
        IRepository InitVBAProject(string directory);

        [DispId(6)]
        [Description("Pushes commits in the CurrentBranch of the Local repo to the Remote.")]
        void Push();

        [DispId(7)]
        [Description("Fetches the specified remote for tracking.\n If argument is not supplied, returns a default remote defined by implementation.")]
        void Fetch([Optional] string remoteName);

        [DispId(8)]
        [Description("Fetches the currently tracking remote and merges it into the CurrentBranch.")]
        void Pull();

        [DispId(9)]
        [Description("Stages and Commits all modified files to the CurrentBranch.")]
        void Commit(string message);

        [DispId(10)]
        [Description("Merges the source branch into the desitnation.")]
        void Merge(string sourceBranch, string destinationBranch);

        [DispId(11)]
        [Description("Checks out the target branch.")]
        void Checkout(string branch);

        [DispId(12)]
        [Description("Creates and checks out a new branch.")]
        void CreateBranch(string branch);

        [DispId(18)]
        [Description("Deletes the specified branch from the local repository.")]
        void DeleteBranch(string branch);

        [DispId(13)]
        [Description("Undoes uncommitted changes to a particular file.")]
        void Undo(string filePath);

        [DispId(14)]
        [Description("Reverts entire branch to the last commit.")]
        void Revert();

        [DispId(15)]
        [Description("Adds untracked file to repository.")]
        void AddFile(string filePath);

        [DispId(16)]
        [Description("Removes file from tracking.")]
        void RemoveFile(string filePath, bool removeFromWorkingDirectory);

        [DispId(17)]
        [Description("Returns a collection of file status entries.\n Semantically the same as calling $git status.")]
        IEnumerable Status();
    }
}
