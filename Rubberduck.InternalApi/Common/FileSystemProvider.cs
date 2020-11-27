using System.IO.Abstractions;

namespace Rubberduck.InternalApi.Common
{
    public static class FileSystemProvider
    {
        static FileSystemProvider()
        {
            FileSystem = new FileSystem();
        }

        /// <remarks>The property is injectable for use in unit testing. Requires use of <see cref="System.Runtime.CompilerServices.InternalsVisibleToAttribute"/>.</remarks>
        public static IFileSystem FileSystem { get; internal set; }
    }
}
