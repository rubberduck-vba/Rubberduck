using System.IO;

namespace Rubberduck.VBEditor.Utility
{
    public interface IFileExistenceChecker
    {
        bool FileExists(string filename);
    }

    public class FileExistenceChecker : IFileExistenceChecker
    {
        public bool FileExists(string filename)
        {
            return File.Exists(filename);
        }
    }
}