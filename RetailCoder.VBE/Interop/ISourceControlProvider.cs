using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;
using System.Runtime.InteropServices;

namespace Rubberduck.Interop
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
        IRepository Clone(string remotePathOrUrl, string workingDirectory);

        [DispId(4)]
        IRepository Init(string directory, bool bare = false);

        [DispId(5)]
        IRepository InitVBAProject(string directory);

        [DispId(6)]
        void Push();

        [DispId(7)]
        void Fetch([Optional] string remoteName);

        [DispId(8)]
        void Pull();

        [DispId(9)]
        void Commit(string message);

        [DispId(10)]
        void Merge(string sourceBranch, string destinationBranch);

        [DispId(11)]
        void Checkout(string branch);

        [DispId(12)]
        void CreateBranch(string branch);

        [DispId(13)]
        void Undo(string filePath);

        [DispId(14)]
        void Revert();

        [DispId(15)]
        void AddFile(string filePath);

        [DispId(16)]
        void RemoveFile(string filePath);

        [DispId(17)]
        IFileStatusEntries Status();
    }
}
