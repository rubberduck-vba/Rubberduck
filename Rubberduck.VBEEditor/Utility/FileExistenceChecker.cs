using Rubberduck.InternalApi.Common;
using System.IO.Abstractions;

namespace Rubberduck.VBEditor.Utility
{
    public interface IFileExistenceChecker
    {
        bool FileExists(string filename);
    }

    public class FileExistenceChecker : IFileExistenceChecker
    {
        private readonly IFileSystem _fileSystem;

        public FileExistenceChecker(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        public bool FileExists(string filename)
        {
            return _fileSystem.File.Exists(filename);
        }
    }
}