using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.SourceControl
{
    public interface IFileStatusEntry
    {
        string FilePath { get; }
        LibGit2Sharp.FileStatus FileStatus { get; }
    }

    public class FileStatusEntry : IFileStatusEntry
    {
        public string FilePath { get; private set; }
        public LibGit2Sharp.FileStatus FileStatus { get; private set; }

        public FileStatusEntry(string filePath, LibGit2Sharp.FileStatus fileStatus)
        {
            this.FilePath = filePath;
            this.FileStatus = fileStatus;
        }

        public FileStatusEntry(LibGit2Sharp.StatusEntry status)
            : this(status.FilePath, status.State) { }
    }
}
