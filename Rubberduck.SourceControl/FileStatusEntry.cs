using System.Runtime.InteropServices;
using LibGit2Sharp;

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
        FileStatus FileStatus { get; }
    }

    [ComVisible(true)]
    [Guid("13AA3AF6-1397-4017-9E97-CBAD6A65FAFA")]
    [ProgId("Rubberduck.FileStatus")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class FileStatusEntry : IFileStatusEntry
    {
        public string FilePath { get; private set; }
        public FileStatus FileStatus { get; private set; }

        private FileStatusEntry(string filePath)
        {
            FilePath = filePath;
        }

        public FileStatusEntry(string filePath, LibGit2Sharp.FileStatus fileStatus)
            :this(filePath)
        {
            FileStatus = (FileStatus)fileStatus;
        }

        public FileStatusEntry(string filePath, FileStatus fileStatus)
            :this(filePath)
        {
            FileStatus = fileStatus;
        }

        public FileStatusEntry(StatusEntry status)
            : this(status.FilePath, status.State) { }
    }
}
