using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace Rubberduck.SourceControl
{
    [ComVisible(true)]
    [Guid("577CB2D3-A84B-44FF-94EF-F4FC78363D74")]
    public interface IFileStatusEntry
    {
        [DispId(0)]
        string FilePath { get; }

        //todo: find a way to make this com visible, even if you have to borrow the source code and cast (int) between them.
        [DispId(1)]
        LibGit2Sharp.FileStatus FileStatus { get; }
    }

    [ComVisible(true)]
    [Guid("13AA3AF6-1397-4017-9E97-CBAD6A65FAFA")]
    [ProgId("Rubberduck.FileStatus")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
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
