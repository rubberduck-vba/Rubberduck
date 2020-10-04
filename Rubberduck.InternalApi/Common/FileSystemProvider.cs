using System.IO.Abstractions;

namespace Rubberduck.InternalApi.Common
{
    public static class FileSystemProvider
    {
        static FileSystemProvider()
        {
            FileSystem = new FileSystem();
        }

        public static IFileSystem FileSystem { get; internal set; }
    }
}
