using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.SourceControl
{
    public interface ISourceControlProvider
    {
        Repository CurrentRepository { get; set; }
        string CurrentBranch { get; }
        IEnumerable<string> Branches { get; }

        /// <summary>Clone a remote repository.</summary>
        /// <param name="remotePathOrUrl">Either a Url "https://github.com/retailcoder/Rubberduck.git" or a UNC path. "//server/share/path/to/repo.git"</param>
        /// <param name="workingDirectory">Directory the repository will be cloned to.</param>
        /// <returns>Newly cloned repository.</returns>
        Repository Clone(string remotePathOrUrl, string workingDirectory);

        /// <summary>
        /// Creates a new repository in/from the given directory.
        /// </summary>
        /// <param name="directory">The directory where the new repository will be created.</param>
        /// <returns>Newly created repository.</returns>
        Repository Init(string directory, Microsoft.Vbe.Interop.VBProject project);

        void Push();
        void Fetch();
        void Pull();
        void Commit(string message);
        void Merge(string sourceBranch, string destinationBranch);
        void Checkout(string branch);
        void Undo(string filePath);
        void Revert();
        void AddFile(string filePath);
        void RemoveFile(string filePath);
    }
}
